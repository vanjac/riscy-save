#include <common.h>
#include <riscyconst.h>
#include <COMUtils.h>
#include <stdlib.h>
#include <Windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
#include <Ole2.h>
#include <shellapi.h>
#include <ShlObj.h>
#include <shobjidl_core.h>
#include <atlbase.h>

#pragma comment(lib, "User32.lib")
#pragma comment(lib, "Gdi32.lib")
#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Shell32.lib")

/*
TODO:
- Get file name including extension (even if hidden)
    - Detect when file name/ext changes
    - use file name text or combobox dropdown to determine extension
- work with any scaling mode (none, system, per monitor v1/2...)

Application bugs:
- Glitchy drag and drop in Firefox
*/

const wchar_t DROP_BOX_CLASS[] = L"Riscy Save Box";
// undocumented message to get pointer to IShellBrowser
const UINT WM_GETISHELLBROWSER = WM_USER + 7;

static HMODULE gModule;
// custom messages (can't use WM_USER or WM_APP)
const UINT MSG_FILE_DIALOG_READY = RegisterWindowMessage(L"Riscy Save Dialog Ready");
// custom timers (hope they don't conflict!)
const UINT TIMER_SELECT_DROPPED = 0x88880001;
const UINT TIMER_CONFIRM_SELECTION = 0x88880002;
const UINT TIMER_CONFIRM_NAME = 0x88880003;

struct DropBox : IUnknownImpl, IDropTarget, IExplorerBrowserEvents {
    HWND wnd = nullptr;
    HWND fileDialog = nullptr;
    bool fileDialogShown = false;
    CComPtr<IShellItemArray> droppedItems;
    DWORD adviseCookie = 0;
    UINT startTimerOnNavComplete = 0;
    CComPtr<IDropTargetHelper> dropTargetHelper;

    HICON fakeFileIcon;
    POINT iconPos;

    DropBox(HWND fileDialog)
        : fileDialog(fileDialog) {}

    void create();
    CComPtr<IExplorerBrowser> getBrowser();

    // drag to box:
    void droppedDataObject(IDataObject *dataObject);
    void selectDroppedItems();
    void confirmSelection();
    // drag from box:
    void dragFakeFile(POINT cursor);
    void draggedFakeFile(wchar_t *path);
    void confirmName();

    // IUnknown
    STDMETHODIMP_(ULONG) AddRef() override { return IUnknownImpl::AddRef(); }
    STDMETHODIMP_(ULONG) Release() override { return IUnknownImpl::Release(); }
    STDMETHODIMP QueryInterface(REFIID id, void **obj) override {
        static const QITAB interfaces[] = {
            QITABENT(DropBox, IDropTarget),
            QITABENT(DropBox, IExplorerBrowserEvents),
            {},
        };
        return QISearch(this, interfaces, id, obj);
    }

    // IDropTarget
    STDMETHODIMP DragEnter(IDataObject *dataObject, DWORD, POINTL point, DWORD *effect) override {
        *effect &= DROPEFFECT_LINK;
        dropTargetHelper->DragEnter(wnd, dataObject, (POINT *)&point, *effect);
        return S_OK;
    }
    STDMETHODIMP DragOver(DWORD, POINTL point, DWORD *effect) override {
        *effect &= DROPEFFECT_LINK;
        dropTargetHelper->DragOver((POINT *)&point, *effect);
        return S_OK;
    }
    STDMETHODIMP DragLeave() override {
        dropTargetHelper->DragLeave();
        return S_OK;
    }
    STDMETHODIMP Drop(IDataObject *dataObject, DWORD, POINTL point, DWORD *effect) override {
        // https://devblogs.microsoft.com/oldnewthing/20100503-00/?p=14183
        // https://devblogs.microsoft.com/oldnewthing/20130204-00/?p=5363
        *effect &= DROPEFFECT_LINK;
        dropTargetHelper->Drop(dataObject, (POINT *)&point, *effect);
        droppedDataObject(dataObject);
        return S_OK;
    }

    // IExplorerBrowserEvents
    STDMETHODIMP OnNavigationPending(PCIDLIST_ABSOLUTE) override { return S_OK; }
    STDMETHODIMP OnViewCreated(IShellView *) override { return S_OK; }
    STDMETHODIMP OnNavigationComplete(PCIDLIST_ABSOLUTE) override {
        if (startTimerOnNavComplete)
            SetTimer(fileDialog, startTimerOnNavComplete, 100, nullptr);
        startTimerOnNavComplete = 0;
        if (CComPtr<IExplorerBrowser> browser = getBrowser())
            browser->Unadvise(adviseCookie);
        return S_OK;
    }
    STDMETHODIMP OnNavigationFailed(PCIDLIST_ABSOLUTE) override { return S_OK; }
};

LRESULT CALLBACK fileDialogSubclassProc(HWND wnd, UINT message,
        WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR refData) {
    DropBox *self = (DropBox *)refData;
    switch (message) {
        case WM_WINDOWPOSCHANGED: {
            WINDOWPOS *windowPos = (WINDOWPOS *)lParam;
            RECT boxRect = {};
            GetWindowRect(self->wnd, &boxRect);
            SetWindowPos(self->wnd, nullptr,
                windowPos->x - (boxRect.right - boxRect.left), windowPos->y, 0, 0,
                SWP_NOACTIVATE | SWP_NOSIZE | SWP_NOZORDER);

            if ((windowPos->flags & SWP_SHOWWINDOW) && !self->fileDialogShown) {
                // https://devblogs.microsoft.com/oldnewthing/20060925-02/?p=29603
                self->fileDialogShown = true;
                PostMessage(wnd, MSG_FILE_DIALOG_READY, 0, 0);
            }
            break;
        }
        case WM_TIMER:
            switch (wParam) {
                case TIMER_SELECT_DROPPED:
                    KillTimer(wnd, wParam);
                    self->selectDroppedItems();
                    return 0;
                case TIMER_CONFIRM_SELECTION:
                    KillTimer(wnd, wParam);
                    self->confirmSelection();
                    return 0;
                case TIMER_CONFIRM_NAME:
                    KillTimer(wnd, wParam);
                    self->confirmName();
                    return 0;
            }
            break;
    }
    if (message == MSG_FILE_DIALOG_READY) {
        ShowWindow(self->wnd, SW_SHOWNORMAL);
        return 0;
    }
    return DefSubclassProc(wnd, message, wParam, lParam);
}

LRESULT CALLBACK dropBoxProc(HWND wnd, UINT message, WPARAM wParam, LPARAM lParam) {
    DropBox *self = (DropBox *)GetWindowLongPtr(wnd, GWLP_USERDATA);
    switch (message) {
        case WM_NCCREATE: {
            CREATESTRUCT *create = (CREATESTRUCT *)lParam;
            self = (DropBox *)create->lpCreateParams;
            SetWindowLongPtr(wnd, GWLP_USERDATA, (LONG_PTR)self);
            self->wnd = wnd;
            self->AddRef(); // keep alive while window open
            break;
        }
        case WM_NCDESTROY:
            self->Release();
            break;
        case WM_CREATE: {
            SetWindowSubclass(self->fileDialog, fileDialogSubclassProc, 0, (DWORD_PTR)self);
            RegisterDragDrop(wnd, self);
            self->dropTargetHelper.CoCreateInstance(CLSID_DragDropHelper);

            SHSTOCKICONINFO iconInfo = {sizeof(iconInfo)};
            SHGetStockIconInfo(SIID_DOCNOASSOC, SHGSI_ICON | SHGSI_LARGEICON | SHGSI_SHELLICONSIZE,
                &iconInfo);
            self->fakeFileIcon = iconInfo.hIcon;
            HIMAGELIST largeIml = nullptr, smallIml = nullptr;
            Shell_GetImageLists(&largeIml, &smallIml);
            int iconWidth, iconHeight;
            ImageList_GetIconSize(largeIml, &iconWidth, &iconHeight);
            CREATESTRUCT *create = (CREATESTRUCT *)lParam;
            self->iconPos.x = (create->cx - iconWidth) / 2;
            self->iconPos.y = (create->cy - iconHeight) / 2;
            break;
        }
        case WM_DESTROY:
            RemoveWindowSubclass(self->fileDialog, fileDialogSubclassProc, 0);
            RevokeDragDrop(wnd);
            DestroyIcon(self->fakeFileIcon);
            break;
        case WM_CLOSE:
            return 0; // no
        case WM_MOUSEACTIVATE:
            if (LOWORD(lParam) == HTCLIENT && HIWORD(lParam) == WM_LBUTTONDOWN)
                return MA_NOACTIVATE; // allow dragging without activating
            break;
        case WM_LBUTTONDOWN: {
            POINT cursor = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
            if (DragDetect(wnd, cursor)) {
                self->dragFakeFile(cursor);
            } else {
                SetActiveWindow(wnd);
            }
            break;
        }
        case WM_PAINT: {
            PAINTSTRUCT paint;
            BeginPaint(wnd, &paint);
            DrawIconEx(paint.hdc, self->iconPos.x, self->iconPos.y, self->fakeFileIcon,
                0, 0, 0, nullptr, DI_NORMAL);
            EndPaint(wnd, &paint);
        }
    }
    return DefWindowProc(wnd, message, wParam, lParam);
}

void DropBox::create() {
    CreateWindow(DROP_BOX_CLASS, nullptr, WS_POPUP | WS_DLGFRAME,
        0, 0, 120, 120, fileDialog, nullptr, gModule, this);
}

CComPtr<IExplorerBrowser> DropBox::getBrowser() {
    IShellBrowser *shellBrowser =
        (IShellBrowser *)SendMessage(fileDialog, WM_GETISHELLBROWSER, 0, 0);
    CComQIPtr<IExplorerBrowser> explorerBrowser(shellBrowser);
    return explorerBrowser;
}

void DropBox::droppedDataObject(IDataObject *dataObject) {
    CComPtr<IExplorerBrowser> browser = getBrowser();
    if (!browser) {
        debugOut(L"Couldn't get browser\n");
        return;
    }

    if (FAILED(SHCreateShellItemArrayFromDataObject(dataObject, IID_PPV_ARGS(&droppedItems))))
        return;
    CComPtr<IShellItem> firstItem;
    if (FAILED(droppedItems->GetItemAt(0, &firstItem)))
        return;
    CComPtr<IShellItem> parentItem;
    if (FAILED(firstItem->GetParent(&parentItem)))
        return;

    startTimerOnNavComplete = TIMER_SELECT_DROPPED;
    browser->Advise(this, &adviseCookie);
    browser->BrowseToObject(parentItem, SBSP_ABSOLUTE);
}

void DropBox::selectDroppedItems() {
    if (!droppedItems)
        return;
    CComPtr<IExplorerBrowser> browser = getBrowser();
    if (!browser)
        return;
    CComPtr<IShellView> shellView;
    if (FAILED(browser->GetCurrentView(IID_PPV_ARGS(&shellView))))
        return;
    CComPtr<IEnumShellItems> enumItems;
    droppedItems->EnumItems(&enumItems);
    droppedItems = nullptr;
    if (!enumItems)
        return;

    shellView->SelectItem(nullptr, SVSI_DESELECTOTHERS);
    for (CComPtr<IShellItem> item; enumItems->Next(1, &item, nullptr) == S_OK; item.Release()) {
        CComHeapPtr<wchar_t> name;
        if (SUCCEEDED(item->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &name))) {
            debugOut(L"Got file %s\n", &*name);
        }
        CComHeapPtr<ITEMID_CHILD> itemIDChild;
        if (SUCCEEDED(CComQIPtr<IParentAndItem>(item)
                ->GetParentAndItem(nullptr, nullptr, &itemIDChild))) {
            shellView->SelectItem(itemIDChild, SVSI_SELECT | SVSI_FOCUSED | SVSI_ENSUREVISIBLE);
        }
    }

    SetTimer(fileDialog, TIMER_CONFIRM_SELECTION, 100, nullptr);
}

void DropBox::confirmSelection() {
    CComPtr<IExplorerBrowser> browser = getBrowser();
    if (!browser)
        return;
    CComPtr<IFolderView2> folderView;
    if (SUCCEEDED(browser->GetCurrentView(IID_PPV_ARGS(&folderView))))
        folderView->InvokeVerbOnSelection(nullptr); // TODO does this work for folders?
}

void DropBox::dragFakeFile(POINT cursor) {
    wchar_t tempPath[MAX_PATH];
    DWORD pathLen = GetTempPath(_countof(tempPath), tempPath);
    if (pathLen + _countof(TEMP_FOLDER_NAME) + 1 >= _countof(tempPath))
        return;
    CopyMemory(tempPath + pathLen, TEMP_FOLDER_NAME, sizeof(TEMP_FOLDER_NAME));
    int err = SHCreateDirectoryEx(nullptr, tempPath, nullptr); // TODO use a namespace extension?
    if (err != ERROR_SUCCESS && err != ERROR_FILE_EXISTS && err != ERROR_ALREADY_EXISTS)
        return;
    CComPtr<IShellItem> tempItem;
    if (FAILED(SHCreateItemFromParsingName(tempPath, nullptr, IID_PPV_ARGS(&tempItem))))
        return;
    CComPtr<IDataObject> dataObject;
    if (FAILED(tempItem->BindToHandler(nullptr, BHID_SFUIObject, IID_PPV_ARGS(&dataObject))))
        return;

    // set drag image...
    CComPtr<IDragSourceHelper> dragHelper;
    if (SUCCEEDED(dragHelper.CoCreateInstance(CLSID_DragDropHelper))) {
        ICONINFO iconInfo = {};
        GetIconInfo(fakeFileIcon, &iconInfo);
        BITMAP colorBitmap = {};
        GetObject(iconInfo.hbmColor, sizeof(colorBitmap), &colorBitmap);
        SHDRAGIMAGE dragImage = {};
        dragImage.sizeDragImage = {colorBitmap.bmWidth, colorBitmap.bmHeight};
        dragImage.ptOffset = {cursor.x - iconPos.x, cursor.y - iconPos.y};
        dragImage.hbmpDragImage = iconInfo.hbmColor;
        dragHelper->InitializeFromBitmap(&dragImage, dataObject);
        DeleteBitmap(iconInfo.hbmColor);
        DeleteBitmap(iconInfo.hbmMask);
    }

    HANDLE mailSlot = CreateMailslot(MAIL_SLOT_NAME, 0, 1000, nullptr);
    if (mailSlot == INVALID_HANDLE_VALUE) {
        debugOut(L"Couldn't create mailslot: %d\n", GetLastError());
        return;
    }
    DWORD effect;
    if (SHDoDragDrop(nullptr, dataObject, nullptr, DROPEFFECT_MOVE, &effect) == DRAGDROP_S_DROP) {
        OVERLAPPED overlap = {};
        overlap.hEvent = CreateEvent(nullptr, FALSE, FALSE, nullptr);
        char message[424]; // max length of mailslot message
        DWORD bytesRead;
        if (!ReadFile(mailSlot, message, sizeof(message), &bytesRead, &overlap)) {
            debugOut(L"Didn't receive response\n"); // TODO: try again?
            MessageBeep(MB_OK);
        } else {
            wchar_t *path = (wchar_t *)message;
            wchar_t *end = path + lstrlen(path) - _countof(TEMP_FOLDER_NAME); // parent dir
            if (end >= path) {
                *end = 0;
                debugOut(L"Received path: %s\n", (wchar_t *)path);
                draggedFakeFile(path);
            }
        }
        CloseHandle(overlap.hEvent);
    }
    CloseHandle(mailSlot);
}

void DropBox::draggedFakeFile(wchar_t *path) {
    CComPtr<IExplorerBrowser> browser = getBrowser();
    if (!browser) {
        debugOut(L"Couldn't get browser\n");
        return;
    }
    CComPtr<IShellItem> item;
    if (FAILED(SHCreateItemFromParsingName(path, nullptr, IID_PPV_ARGS(&item))))
        return;
    startTimerOnNavComplete = TIMER_CONFIRM_NAME;
    browser->Advise(this, &adviseCookie);
    browser->BrowseToObject(item, SBSP_ABSOLUTE);
}

void DropBox::confirmName() {
    // TODO make sure to deselect first in case the previous folder is selected
    PostMessage(fileDialog, WM_COMMAND, MAKEWPARAM(IDOK, BN_CLICKED),
        (LPARAM)GetDlgItem(fileDialog, IDOK));
}


HWND findFileDialog(HWND wnd, CREATESTRUCT *create) {
    if (!create->hwndParent)
        return nullptr;
    wchar_t className[256];
    GetClassName(wnd, className, _countof(className));
    if (lstrcmp(className, L"SHELLDLL_DefView") != 0)
        return nullptr;
    GetClassName(create->hwndParent, className, _countof(className));
    if (lstrcmp(className, L"ExplorerBrowserControl") == 0)
        return nullptr;
    HWND root = GetAncestor(create->hwndParent, GA_ROOT);
    if (MAKEINTATOM(GetClassWord(root, GCW_ATOM)) != WC_DIALOG)
        return nullptr;
    DWORD_PTR refData;
    if (GetWindowSubclass(root, fileDialogSubclassProc, 0, &refData)) // already subclassed
        return nullptr;
    return root;
}

/* EXPORTED */

LRESULT CALLBACK cbtHookProc(int code, WPARAM wParam, LPARAM lParam) {
    if (code == HCBT_CREATEWND) {
        HWND wnd = (HWND)wParam;
        CBT_CREATEWND *createWnd = (CBT_CREATEWND *)lParam;
        if (HWND dlg = findFileDialog(wnd, createWnd->lpcs)) {
            debugOut(L"File dialog created!\n");
            // TODO: reposition dialog so dropbox is visible
            CComPtr<DropBox> dropBox;
            dropBox.Attach(new DropBox(dlg));
            dropBox->create();
        }
    }
    return CallNextHookEx(nullptr, code, wParam, lParam );
}

BOOL WINAPI DllMain(HINSTANCE module, DWORD reason, LPVOID) {
    switch (reason) {
        case DLL_PROCESS_ATTACH: {
            gModule = module;
            DisableThreadLibraryCalls(module);

            WNDCLASS wndClass = {};
            wndClass.lpszClassName = DROP_BOX_CLASS;
            wndClass.lpfnWndProc = dropBoxProc;
            wndClass.hCursor = LoadCursor(NULL, IDC_ARROW);
            wndClass.hInstance = module;
            RegisterClass(&wndClass);
            break;
        }
        case DLL_PROCESS_DETACH:
            UnregisterClass(DROP_BOX_CLASS, module);
            break;
    }
    return TRUE;
}
