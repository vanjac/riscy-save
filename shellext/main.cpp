#include <common.h>
#include <riscyconst.h>
#include <COMUtils.h>
#include <Windows.h>
#include <ShlObj.h>
#include <ShObjIdl.h>
#include <Shlwapi.h>
#include <shellapi.h>

#pragma comment(lib, "Shlwapi.lib")

static LONG lockCount = 0;

// {bd515f55-a1f5-42c2-b245-dd40fd0cf9b0}
const CLSID CLSID_RiscyExtension =
    {0xbd515f55, 0xa1f5, 0x42c2, {0xb2, 0x45, 0xdd, 0x40, 0xfd, 0x0c, 0xf9, 0xb0}};
class RiscyExtension : public IUnknownImpl, ICopyHook {
public:
    void sendDragResult(const wchar_t *path);

    // IUnknown
    STDMETHODIMP_(ULONG) AddRef() override { return IUnknownImpl::AddRef(); };
    STDMETHODIMP_(ULONG) Release() override { return IUnknownImpl::Release(); };
    STDMETHODIMP QueryInterface(REFIID id, void **obj) override {
        static const QITAB interfaces[] = {
            QITABENT(RiscyExtension, ICopyHook),
            {},
        };
        return QISearch(this, interfaces, id, obj);
    }

    // ICopyHook
    STDMETHODIMP_(UINT) CopyCallback(HWND, UINT func, UINT,
            const wchar_t *srcFile, DWORD,
            const wchar_t *destFile, DWORD) override {
        if (func != FO_MOVE && func != FO_COPY)
            return IDYES;
        const wchar_t *end = srcFile + lstrlen(srcFile) - _countof(TEMP_FOLDER_NAME) + 1;
        if (end >= srcFile && lstrcmpi(end, TEMP_FOLDER_NAME) == 0) {
            debugOut(L"Moved temp folder to %s\n", destFile);
            sendDragResult(destFile);
            return IDCANCEL;
        }
        return IDYES;
    }
};

class RiscyExtensionFactory : public IClassFactory {
public:
    // IUnknown
    STDMETHODIMP_(ULONG) AddRef() override { return 2; }
    STDMETHODIMP_(ULONG) Release() override { return 1; }
    STDMETHODIMP QueryInterface(REFIID id, void **obj) override {
        static const QITAB interfaces[] = {
            QITABENT(RiscyExtensionFactory, IClassFactory),
            {},
        };
        return QISearch(this, interfaces, id, obj);
    }

    // IClassFactory
    STDMETHODIMP CreateInstance(IUnknown *outer, REFIID id, void **obj) override {
        *obj = nullptr;
        if (outer)
            return CLASS_E_NOAGGREGATION;
        RiscyExtension *ext = new RiscyExtension();
        HRESULT hr = ext->QueryInterface(id, obj);
        ext->Release();
        return hr;
    }
    STDMETHODIMP LockServer(BOOL lock) override {
        if (lock)
            InterlockedIncrement(&lockCount);
        else
            InterlockedDecrement(&lockCount);
        return S_OK;
    }
};

static RiscyExtensionFactory factory;

void RiscyExtension::sendDragResult(const wchar_t *path) {
    HANDLE mailSlotFile = CreateFile(MAIL_SLOT_NAME, GENERIC_WRITE, FILE_SHARE_READ, nullptr,
        OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (mailSlotFile == INVALID_HANDLE_VALUE) {
        debugOut(L"Couldn't open mail slot: %d\n", GetLastError());
        return;
    }
    DWORD bytesWritten;
    if (!WriteFile(mailSlotFile, path, (lstrlen(path) + 1) * sizeof(wchar_t),
            &bytesWritten, nullptr)) {
        debugOut(L"Error writing to mail slot: %d\n", GetLastError());
    }
    CloseHandle(mailSlotFile);
}

/* EXPORTED */

BOOL WINAPI DllMain(HINSTANCE module, DWORD reason, LPVOID) {
    switch (reason) {
        case DLL_PROCESS_ATTACH:
            DisableThreadLibraryCalls(module);
            break;
    }
    return TRUE;
}

STDAPI DllGetClassObject(REFCLSID clsid, REFIID id, LPVOID *obj) {
    if (clsid == CLSID_RiscyExtension)
        return factory.QueryInterface(id, obj);
    *obj = nullptr;
    return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow() {
    return lockCount ? S_FALSE : S_OK; // mistake in oldnewthing article?
}