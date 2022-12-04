#include <common.h>
#include <Windows.h>
#include <ShlObj_core.h>
#include <shellapi.h>
#include <CommCtrl.h>

#pragma comment(lib, "User32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Comctl32.lib")

LRESULT CALLBACK notifyWindowProc(HWND wnd, UINT message, WPARAM wParam, LPARAM lParam) {
    if (message == WM_USER) {
        if (LOWORD(lParam) == WM_CONTEXTMENU)
            PostQuitMessage(0);
        return 0;
    }
    return DefWindowProc(wnd, message, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int) {
#ifdef RISCY_MEMLEAKS
    _CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF);
    debugOut(L"Compiled with memory leak detection\n");
#endif
    InitCommonControls();

#ifdef _WIN64
    wchar_t exePath[MAX_PATH];
    PROCESS_INFORMATION process32Info = {};
    if (DWORD exePathLen = GetModuleFileName(instance, exePath, MAX_PATH)) {
        CopyMemory(exePath + exePathLen - 6, L"32", 2*sizeof(wchar_t)); // TODO no
        STARTUPINFO startup = {sizeof(startup)};
        debugOut(L"Starting process %s\n", exePath);
        CreateProcess(exePath, L"riscysave32.exe", nullptr, nullptr, FALSE,
            DETACHED_PROCESS, nullptr, nullptr, &startup, &process32Info);
    }

    HINSTANCE libModule = LoadLibrary(L"riscylib64.dll");
#else
    HINSTANCE libModule = LoadLibrary(L"riscylib32.dll");
#endif
    if (!libModule) {
        debugOut(L"unable to load module\n");
        return 0;
    }
    HOOKPROC hookProc = (HOOKPROC)GetProcAddress(libModule, "cbtHookProc");
    HHOOK hook = SetWindowsHookEx(WH_CBT, hookProc, libModule, 0); // :)
    if (!hook) {
        debugOut(L"couldn't create hook %d\n", GetLastError());
        return 0;
    }


    // https://learn.microsoft.com/en-us/windows/win32/shell/handlers#registering-shell-extension-handlers
    SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nullptr, nullptr);


    WNDCLASS notifyClass = {};
    notifyClass.lpszClassName = L"Riscy Save";
    notifyClass.lpfnWndProc = notifyWindowProc;
    notifyClass.hInstance = instance;
    RegisterClass(&notifyClass);


    HWND messageWindow = CreateWindow(notifyClass.lpszClassName, nullptr, 0,
        0, 0, 0, 0, HWND_MESSAGE, nullptr, instance, 0);
    if (!messageWindow)
        return 0;

    NOTIFYICONDATA notify = {sizeof(notify)};
    notify.uFlags = NIF_ICON | NIF_TIP | NIF_SHOWTIP | NIF_MESSAGE;
    notify.uCallbackMessage = WM_USER;
    notify.hWnd = messageWindow;
    notify.uVersion = NOTIFYICON_VERSION_4;
    notify.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
#ifdef _WIN64
    wchar_t title[] = L"Riscy Save 64";
#else
    wchar_t title[] = L"Riscy Save 32";
#endif
    CopyMemory(notify.szTip, title, sizeof(title));
    if (!Shell_NotifyIcon(NIM_ADD, &notify)) {
        debugOut(L"Couldn't add notification icon\n");
        return 0;
    }
    Shell_NotifyIcon(NIM_SETVERSION, &notify);

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    Shell_NotifyIcon(NIM_DELETE, &notify);

#ifdef _WIN64
    if (process32Info.hProcess) {
        PostThreadMessage(process32Info.dwThreadId, WM_QUIT, 0, 0);
        CloseHandle(process32Info.hProcess);
        CloseHandle(process32Info.hThread);
    }
#endif


    if (!UnhookWindowsHookEx(hook))
        debugOut(L"Couldn't unhook\n");
    // https://stackoverflow.com/a/4076495/11525734
    DWORD_PTR result;
    SendMessageTimeout(HWND_BROADCAST, WM_NULL, 0, 0,
        SMTO_ABORTIFHUNG|SMTO_NOTIMEOUTIFNOTHUNG, 1000, &result);

    return (int)msg.wParam;
}
