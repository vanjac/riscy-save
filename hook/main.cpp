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
- work with any scaling mode (none, system, per monitor v1/2...)
- automatically change filter to All Files before opening/saving (enter * in name box)

Application bugs:
- Glitchy drag and drop in Firefox
*/

const wchar_t DROP_BOX_CLASS[] = L"Riscy Save Box";
// undocumented message to get pointer to IShellBrowser
const UINT WM_GETISHELLBROWSER = WM_USER + 7;
// dialog control ids
const int DLG_VISTA_SAVE_FILENAME = 1001; // Edit
const int DLG_FILENAME_COMBO = 1148;
const int DLG_FILENAME_EDIT = 1152; // win95 style dialogs without a combo box

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
    HWND fileNameCtl = nullptr; // could be Edit or ComboBox
    wchar_t fileName[MAX_PATH];

    CComPtr<IShellItemArray> droppedItems;
    DWORD adviseCookie = 0;
    UINT startTimerOnNavComplete = 0;

    CComPtr<IDropTargetHelper> dropTargetHelper;
    CComPtr<IDataObject> dropObject;

    HICON fakeFileIcon;
    RECT iconRect;

    DropBox(HWND fileDialog)
        : fileDialog(fileDialog) {}

    void create();
    CComPtr<IExplorerBrowser> getBrowser();
    void updateFileName();
    void updateDrag(DWORD effect);

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
        dropObject = dataObject;
        *effect &= DROPEFFECT_LINK;
        updateDrag(*effect);
        dropTargetHelper->DragEnter(wnd, dataObject, (POINT *)&point, *effect);
        return S_OK;
    }
    STDMETHODIMP DragOver(DWORD, POINTL point, DWORD *effect) override {
        *effect &= DROPEFFECT_LINK;
        updateDrag(*effect);
        dropTargetHelper->DragOver((POINT *)&point, *effect);
        return S_OK;
    }
    STDMETHODIMP DragLeave() override {
        updateDrag(DROPEFFECT_NONE);
        dropTargetHelper->DragLeave();
        dropObject = nullptr;
        return S_OK;
    }
    STDMETHODIMP Drop(IDataObject *dataObject, DWORD, POINTL point, DWORD *effect) override {
        // https://devblogs.microsoft.com/oldnewthing/20100503-00/?p=14183
        // https://devblogs.microsoft.com/oldnewthing/20130204-00/?p=5363
        *effect &= DROPEFFECT_LINK;
        dropTargetHelper->Drop(dataObject, (POINT *)&point, *effect);
        droppedDataObject(dataObject);
        dropObject = nullptr;
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


struct FindDlgItemRecursiveData { int id; HWND result; };

BOOL CALLBACK findDlgItemRecursiveProc(HWND wnd, LPARAM param) {
    FindDlgItemRecursiveData *data = (FindDlgItemRecursiveData *)param;
    if (GetDlgCtrlID(wnd) == data->id) {
        data->result = wnd;
        return FALSE;
    }
    return TRUE;
}

HWND findDlgItemRecursive(HWND root, int id) {
    FindDlgItemRecursiveData data = {id, nullptr};
    EnumChildWindows(root, findDlgItemRecursiveProc, (LPARAM)&data);
    return data.result;
}


LRESULT CALLBACK fileNameEditProc(HWND wnd, UINT message,
        WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR refData) {
    DropBox *self = (DropBox *)refData;
    LRESULT res = DefSubclassProc(wnd, message, wParam, lParam);
    if (message == WM_SETTEXT)
        self->updateFileName();
    return res;
}

LRESULT CALLBACK fileNameComboBoxProc(HWND wnd, UINT message,
        WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR refData) {
    DropBox *self = (DropBox *)refData;
    if (message == WM_COMMAND && (HWND)lParam == self->fileNameCtl && HIWORD(wParam) == EN_CHANGE)
        self->updateFileName();
    return DefSubclassProc(wnd, message, wParam, lParam);
}

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
        case WM_COMMAND: {
            int cmd = HIWORD(wParam);
            if ((HWND)lParam == self->fileNameCtl
                    && (cmd == CBN_EDITCHANGE || cmd == CBN_SELCHANGE || cmd == EN_CHANGE)) {
                self->updateFileName();
            }
            break;
        }
    }
    if (message == MSG_FILE_DIALOG_READY) {
        ShowWindow(self->wnd, SW_SHOWNORMAL);
        self->fileNameCtl = GetDlgItem(self->fileDialog, DLG_FILENAME_COMBO);
        if (!self->fileNameCtl)
            self->fileNameCtl = GetDlgItem(self->fileDialog, DLG_FILENAME_EDIT);
        if (!self->fileNameCtl) {
            self->fileNameCtl = findDlgItemRecursive(self->fileDialog, DLG_VISTA_SAVE_FILENAME);
            SetWindowSubclass(self->fileNameCtl, fileNameEditProc, 0, (DWORD_PTR)self);
            SetWindowSubclass(GetParent(self->fileNameCtl),
                fileNameComboBoxProc, 0, (DWORD_PTR)self);
        }
        if (self->fileNameCtl) {
            self->updateFileName();
        } else {
            debugOut(L"Failed to find filename box!\n");
            self->fileName[0] = 0;
        }
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
            self->iconRect.left = (create->cx - iconWidth) / 2;
            self->iconRect.top = (create->cy - iconHeight) / 4;
            self->iconRect.right = self->iconRect.left + iconWidth;
            self->iconRect.bottom = self->iconRect.top + iconHeight;
            break;
        }
        case WM_DESTROY:
            if (self->fileNameCtl) {
                RemoveWindowSubclass(self->fileNameCtl, fileNameEditProc, 0);
                RemoveWindowSubclass(GetParent(self->fileNameCtl), fileNameComboBoxProc, 0);
            }
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
            if (self->fileName[0] && DragDetect(wnd, cursor)) {
                self->dragFakeFile(cursor);
            } else {
                SetActiveWindow(wnd);
            }
            break;
        }
        case WM_PAINT: {
            PAINTSTRUCT paint;
            BeginPaint(wnd, &paint);
            if (self->fileName[0]) {
                DrawIconEx(paint.hdc, self->iconRect.left, self->iconRect.top, self->fakeFileIcon,
                    0, 0, 0, nullptr, DI_NORMAL);
                RECT clientRect = {};
                GetClientRect(wnd, &clientRect);
                RECT labelRect = {0, self->iconRect.bottom, clientRect.right, clientRect.bottom};
                DrawText(paint.hdc, self->fileName, -1, &labelRect,
                    DT_CENTER | DT_TOP | DT_WORDBREAK | DT_WORD_ELLIPSIS | DT_NOPREFIX);
            }
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

void DropBox::updateFileName() {
    GetWindowText(fileNameCtl, fileName, _countof(fileName));
    if (PathCleanupSpec(nullptr, fileName) & (PCS_REPLACEDCHAR | PCS_REMOVEDCHAR)) {
        fileName[0] = 0; // not a valid filename
    } else {
        SHFILEINFO fileInfo = {};
        if (SHGetFileInfo(fileName, FILE_ATTRIBUTE_NORMAL, &fileInfo, sizeof(fileInfo),
                SHGFI_USEFILEATTRIBUTES | SHGFI_ICON | SHGFI_LARGEICON | SHGFI_SHELLICONSIZE)) {
            DestroyIcon(fakeFileIcon);
            fakeFileIcon = fileInfo.hIcon;
        }
    }
    RedrawWindow(wnd, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE);
}

void DropBox::updateDrag(DWORD effect) {
    DROPDESCRIPTION *dropDesc = (DROPDESCRIPTION *)GlobalAlloc(GPTR, sizeof(DROPDESCRIPTION));
    if (effect & DROPEFFECT_LINK) {
        dropDesc->type = DROPIMAGE_LINK;
        GetWindowText(fileDialog, dropDesc->szMessage, _countof(dropDesc->szMessage));
    } else {
        dropDesc->type = DROPIMAGE_INVALID;
        dropDesc->szMessage[0] = 0;
    }
    dropDesc->szInsert[0] = 0;
    FORMATETC format = { (CLIPFORMAT)RegisterClipboardFormat(CFSTR_DROPDESCRIPTION),
        nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM med = { TYMED_HGLOBAL };
    med.hGlobal = dropDesc;
    dropObject->SetData(&format, &med, true);
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
        dragImage.ptOffset = {cursor.x - iconRect.left, cursor.y - iconRect.top};
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
            wndClass.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
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
