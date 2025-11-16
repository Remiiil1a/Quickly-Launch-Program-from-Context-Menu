#include <windows.h>
#include <commctrl.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <algorithm>
#include <map>

// 显式链接所需的库
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "shlwapi.lib")

// 应用程序结构
struct AppEntry
{
    std::wstring name;        // 注册表项名称
    std::wstring path;        // 程序路径
    std::wstring displayName; // 显示名称
    std::wstring icon;        // 图标路径
    bool isCustom;            // 是否由本程序创建
};

class RightClickManager
{
private:
    std::vector<AppEntry> apps;
    std::vector<AppEntry> allApps; // 存储所有应用，用于过滤
    HWND hMainWindow;
    HWND hListBox;
    HWND hAddButton;
    HWND hRemoveButton;
    HWND hRefreshButton;
    HWND hShowAllCheckbox;
    HANDLE hMutex;
    bool showAllItems; // 是否显示所有项

    // 系统内置的右键菜单项黑名单
    std::vector<std::wstring> systemItems = {
        L"New",
        L"View",
        L"SortBy",
        L"Paste",
        L"PasteShortcut",
        L"DesktopBackground",
        L"Settings",
        L"Display",
        L"GraphicsProperties",
        L"NvDriverUpdate",
        L"Share",
        L"GrantAccess",
        L"PinToQuickAccess",
        L"IncludeInLibrary",
        L"Properties",
        L"Open",
        L"OpenInNewWindow",
        L"Print",
        L"ScanWithMicrosoftDefender"};

public:
    RightClickManager() : hMainWindow(NULL), hListBox(NULL), hAddButton(NULL),
                          hRemoveButton(NULL), hRefreshButton(NULL), hShowAllCheckbox(NULL),
                          hMutex(NULL), showAllItems(false) {}

    ~RightClickManager()
    {
        if (hMutex)
        {
            CloseHandle(hMutex);
        }
    }

    // 检查是否已有实例在运行
    bool IsAlreadyRunning()
    {
        hMutex = CreateMutexW(NULL, TRUE, L"RightClickManager_SingleInstance");
        if (hMutex == NULL)
        {
            return false;
        }

        if (GetLastError() == ERROR_ALREADY_EXISTS)
        {
            return true;
        }

        return false;
    }

    // 激活已运行的实例窗口
    void ActivateExistingInstance()
    {
        HWND hExistingWindow = FindWindowW(L"RightClickManager", L"桌面右键菜单管理器");
        if (hExistingWindow)
        {
            if (IsIconic(hExistingWindow))
            {
                ShowWindow(hExistingWindow, SW_RESTORE);
            }
            SetForegroundWindow(hExistingWindow);
        }
    }

    // 初始化程序
    bool Initialize(HINSTANCE hInstance)
    {
        // 初始化通用控件
        INITCOMMONCONTROLSEX icex;
        icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
        icex.dwICC = ICC_LISTVIEW_CLASSES | ICC_TREEVIEW_CLASSES | ICC_BAR_CLASSES;
        InitCommonControlsEx(&icex);

        // 创建主窗口
        WNDCLASSEXW wc = {};
        wc.cbSize = sizeof(WNDCLASSEXW);
        wc.style = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc = WindowProc;
        wc.hInstance = hInstance;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"RightClickManager";
        wc.hIcon = LoadIcon(hInstance, IDI_APPLICATION);

        RegisterClassExW(&wc);

        // 计算窗口位置使其居中
        int screenWidth = GetSystemMetrics(SM_CXSCREEN);
        int screenHeight = GetSystemMetrics(SM_CYSCREEN);
        int windowWidth = 750;  // 增加宽度以容纳更多信息
        int windowHeight = 450; // 增加高度以容纳新控件
        int x = (screenWidth - windowWidth) / 2;
        int y = (screenHeight - windowHeight) / 2;

        // 修改窗口样式，移除可调整大小的边框
        hMainWindow = CreateWindowExW(
            0,
            L"RightClickManager",
            L"桌面右键菜单管理器",
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
            x, y,
            windowWidth, windowHeight,
            NULL, NULL, hInstance, this);

        if (!hMainWindow)
            return false;

        CreateControls(hInstance);
        LoadAllContextMenuItems();
        ShowWindow(hMainWindow, SW_SHOW);
        UpdateWindow(hMainWindow);

        return true;
    }

    // 创建控件
    void CreateControls(HINSTANCE hInstance)
    {
        // 列表控件 - 添加水平滚动条
        hListBox = CreateWindowExW(
            WS_EX_CLIENTEDGE,
            L"LISTBOX",
            L"",
            WS_CHILD | WS_VISIBLE | LBS_NOTIFY | LBS_HASSTRINGS | WS_VSCROLL | WS_HSCROLL,
            10, 10, 580, 380, // 宽度增加到580
            hMainWindow,
            (HMENU)1001,
            hInstance,
            NULL);

        // 设置列表项高度，使文本更容易阅读
        SendMessage(hListBox, LB_SETITEMHEIGHT, 0, 20);

        // 添加按钮
        hAddButton = CreateWindowW(
            L"BUTTON",
            L"添加程序",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            600, 10, 140, 30, // x=600, 宽度140
            hMainWindow,
            (HMENU)1002,
            hInstance,
            NULL);

        // 删除按钮
        hRemoveButton = CreateWindowW(
            L"BUTTON",
            L"删除选中项",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            600, 50, 140, 30,
            hMainWindow,
            (HMENU)1003,
            hInstance,
            NULL);

        // 刷新按钮
        hRefreshButton = CreateWindowW(
            L"BUTTON",
            L"刷新列表",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            600, 90, 140, 30,
            hMainWindow,
            (HMENU)1004,
            hInstance,
            NULL);

        // 创建复选框 - 这是唯一需要的一个
        hShowAllCheckbox = CreateWindowW(
            L"BUTTON",
            L"非本程序添加项",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
            600, 130, 200, 30, // 增加宽度确保文本完整显示
            hMainWindow,
            (HMENU)1005,
            hInstance,
            NULL);

        // 添加说明文本 - 调整位置避免重叠
        CreateWindowW(
            L"STATIC",
            L"管理所有桌面右键菜单项\n\n"
            L"√ 表示由本程序创建\n"
            L"其他为系统或第三方程序创建\n\n"
            L"提示：可以拖动水平滚动条查看完整路径",
            WS_CHILD | WS_VISIBLE,
            600, 170, 140, 150, // 调整Y位置从180改为170
            hMainWindow,
            NULL,
            hInstance,
            NULL);

        SendMessage(hShowAllCheckbox, BM_SETCHECK, showAllItems ? BST_CHECKED : BST_UNCHECKED, 0);
    }

    // 检查是否为系统项
    bool IsSystemItem(const std::wstring &itemName)
    {
        for (const auto &systemItem : systemItems)
        {
            if (itemName == systemItem)
            {
                return true;
            }
        }
        return false;
    }

    // 加载所有右键菜单项
    void LoadAllContextMenuItems()
    {
        allApps.clear();
        apps.clear();
        SendMessageW(hListBox, LB_RESETCONTENT, 0, 0);

        // 检查桌面右键菜单注册表位置
        HKEY hKey;
        if (RegOpenKeyExW(HKEY_CLASSES_ROOT, L"Directory\\Background\\shell", 0, KEY_READ, &hKey) == ERROR_SUCCESS)
        {
            wchar_t subkeyName[256];
            DWORD index = 0;
            DWORD nameSize = sizeof(subkeyName) / sizeof(wchar_t);

            while (RegEnumKeyExW(hKey, index, subkeyName, &nameSize, NULL, NULL, NULL, NULL) == ERROR_SUCCESS)
            {
                // 跳过系统项
                if (!IsSystemItem(subkeyName))
                {
                    AppEntry app;
                    app.name = subkeyName;
                    app.isCustom = (wcsstr(subkeyName, L"CustomApp_") == subkeyName);

                    // 获取显示名称
                    std::wstring displayPath = L"Directory\\Background\\shell\\";
                    displayPath += subkeyName;

                    HKEY hDisplayKey;
                    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, displayPath.c_str(), 0, KEY_READ, &hDisplayKey) == ERROR_SUCCESS)
                    {
                        wchar_t displayName[256];
                        DWORD nameSize = sizeof(displayName);
                        if (RegQueryValueExW(hDisplayKey, NULL, NULL, NULL, (LPBYTE)displayName, &nameSize) == ERROR_SUCCESS)
                        {
                            app.displayName = displayName;
                        }
                        else
                        {
                            app.displayName = subkeyName; // 如果没有显示名称，使用注册表项名
                        }

                        // 获取图标
                        wchar_t iconPath[1024];
                        DWORD iconSize = sizeof(iconPath);
                        if (RegQueryValueExW(hDisplayKey, L"Icon", NULL, NULL, (LPBYTE)iconPath, &iconSize) == ERROR_SUCCESS)
                        {
                            app.icon = iconPath;
                        }

                        RegCloseKey(hDisplayKey);
                    }
                    else
                    {
                        app.displayName = subkeyName;
                    }

                    // 获取程序路径
                    std::wstring commandPath = displayPath + L"\\command";
                    HKEY hCommandKey;
                    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, commandPath.c_str(), 0, KEY_READ, &hCommandKey) == ERROR_SUCCESS)
                    {
                        wchar_t appPath[1024];
                        DWORD pathSize = sizeof(appPath);
                        if (RegQueryValueExW(hCommandKey, NULL, NULL, NULL, (LPBYTE)appPath, &pathSize) == ERROR_SUCCESS)
                        {
                            app.path = appPath;
                            CleanAppPath(app.path);
                        }
                        RegCloseKey(hCommandKey);
                    }

                    allApps.push_back(app);
                }
                index++;
                nameSize = sizeof(subkeyName) / sizeof(wchar_t);
            }
            RegCloseKey(hKey);
        }

        // 根据显示设置过滤应用列表
        FilterApps();
    }

    // 根据显示设置过滤应用列表
    void FilterApps()
    {
        apps.clear();
        SendMessageW(hListBox, LB_RESETCONTENT, 0, 0);

        for (const auto &app : allApps)
        {
            if (showAllItems || app.isCustom)
            {
                apps.push_back(app);

                // 在列表中显示，标记由本程序创建的项
                std::wstring listText = app.isCustom ? L"√ " : L"  ";
                listText += app.displayName + L" - " + app.path;
                SendMessageW(hListBox, LB_ADDSTRING, 0, (LPARAM)listText.c_str());
            }
        }

        // 设置水平滚动范围，使长文本可以滚动查看
        if (!apps.empty())
        {
            HDC hdc = GetDC(hListBox);
            if (hdc)
            {
                int maxWidth = 0;
                for (const auto &app : apps)
                {
                    std::wstring listText = app.isCustom ? L"√ " : L"  ";
                    listText += app.displayName + L" - " + app.path;

                    SIZE size;
                    GetTextExtentPoint32W(hdc, listText.c_str(), listText.length(), &size);
                    if (size.cx > maxWidth)
                    {
                        maxWidth = size.cx;
                    }
                }
                ReleaseDC(hListBox, hdc);

                // 设置水平滚动范围
                SendMessage(hListBox, LB_SETHORIZONTALEXTENT, maxWidth + 20, 0);
            }
        }
    }

    // 清理应用路径
    void CleanAppPath(std::wstring &path)
    {
        // 移除引号
        if (path.length() >= 2 && path[0] == L'\"' && path[path.length() - 1] == L'\"')
        {
            path = path.substr(1, path.length() - 2);
        }
        // 移除参数
        size_t pos = path.find(L".exe");
        if (pos != std::wstring::npos)
        {
            path = path.substr(0, pos + 4);
        }
    }

    // 添加应用到桌面右键菜单
    bool AddAppToContextMenu(const std::wstring &appPath)
    {
        // 获取程序名称
        size_t lastSlash = appPath.find_last_of(L'\\');
        size_t lastDot = appPath.find_last_of(L'.');
        if (lastSlash == std::wstring::npos)
            return false;

        std::wstring appName = appPath.substr(lastSlash + 1);
        if (lastDot != std::wstring::npos && lastDot > lastSlash)
        {
            appName = appPath.substr(lastSlash + 1, lastDot - lastSlash - 1);
        }

        // 生成唯一的注册表键名
        std::wstring registryKey = L"CustomApp_";
        registryKey += appName;

        // 创建注册表项
        std::wstring shellKey = L"Directory\\Background\\shell\\";
        shellKey += registryKey;

        HKEY hKey;
        LONG result = RegCreateKeyExW(HKEY_CLASSES_ROOT, shellKey.c_str(), 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL);
        if (result == ERROR_SUCCESS)
        {
            // 设置显示名称
            result = RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE *)appName.c_str(), (appName.length() + 1) * sizeof(wchar_t));
            if (result != ERROR_SUCCESS)
            {
                RegCloseKey(hKey);
                return false;
            }

            // 设置图标
            std::wstring iconValue = L"\"";
            iconValue += appPath;
            iconValue += L"\"";

            RegSetValueExW(hKey, L"Icon", 0, REG_SZ, (const BYTE *)iconValue.c_str(), (iconValue.length() + 1) * sizeof(wchar_t));

            RegCloseKey(hKey);
        }
        else
        {
            return false;
        }

        // 创建command子键
        std::wstring commandKey = shellKey + L"\\command";
        if (RegCreateKeyExW(HKEY_CLASSES_ROOT, commandKey.c_str(), 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
        {
            std::wstring commandValue = L"\"";
            commandValue += appPath;
            commandValue += L"\"";

            result = RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE *)commandValue.c_str(), (commandValue.length() + 1) * sizeof(wchar_t));
            RegCloseKey(hKey);

            if (result == ERROR_SUCCESS)
            {
                // 刷新系统，使注册表更改立即生效
                SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

                LoadAllContextMenuItems(); // 刷新列表
                return true;
            }
        }

        return false;
    }

    // 从右键菜单删除应用
    bool RemoveAppFromContextMenu(int index)
    {
        if (index < 0 || index >= (int)apps.size())
            return false;

        AppEntry &app = apps[index];

        // 对于非本程序创建的项，显示额外警告
        if (!app.isCustom)
        {
            std::wstring warningMsg = L"警告：此项不是由本程序创建，可能是系统或其他应用程序的右键菜单项。\n\n";
            warningMsg += L"名称: " + app.displayName + L"\n";
            warningMsg += L"路径: " + app.path + L"\n\n";
            warningMsg += L"确定要删除吗？";

            if (MessageBoxW(hMainWindow, warningMsg.c_str(), L"确认删除系统项", MB_YESNO | MB_ICONWARNING) != IDYES)
            {
                return false;
            }
        }

        std::wstring shellKey = L"Directory\\Background\\shell\\";
        shellKey += app.name;

        // 递归删除注册表项
        if (DeleteRegistryTree(HKEY_CLASSES_ROOT, shellKey.c_str()))
        {
            // 刷新系统
            SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

            LoadAllContextMenuItems(); // 重新加载所有项
            return true;
        }

        return false;
    }

    // 递归删除注册表树
    bool DeleteRegistryTree(HKEY hParentKey, const wchar_t *subkey)
    {
        HKEY hKey;
        if (RegOpenKeyExW(hParentKey, subkey, 0, KEY_READ | KEY_WRITE, &hKey) != ERROR_SUCCESS)
        {
            return false;
        }

        // 枚举并删除所有子键
        wchar_t childKeyName[256];
        DWORD childKeySize = sizeof(childKeyName) / sizeof(wchar_t);

        while (RegEnumKeyExW(hKey, 0, childKeyName, &childKeySize, NULL, NULL, NULL, NULL) == ERROR_SUCCESS)
        {
            std::wstring fullChildKey = subkey;
            fullChildKey += L"\\";
            fullChildKey += childKeyName;

            if (!DeleteRegistryTree(hParentKey, fullChildKey.c_str()))
            {
                RegCloseKey(hKey);
                return false;
            }
            childKeySize = sizeof(childKeyName) / sizeof(wchar_t);
        }

        RegCloseKey(hKey);

        // 删除当前键
        return RegDeleteKeyW(hParentKey, subkey) == ERROR_SUCCESS;
    }

    // 处理按钮点击
    void OnAddButtonClick()
    {
        OPENFILENAMEW ofn;
        wchar_t fileName[MAX_PATH] = L"";

        ZeroMemory(&ofn, sizeof(ofn));
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner = hMainWindow;
        ofn.lpstrFile = fileName;
        ofn.nMaxFile = MAX_PATH;
        ofn.lpstrFilter = L"可执行文件\0*.exe\0所有文件\0*.*\0";
        ofn.nFilterIndex = 1;
        ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

        if (GetOpenFileNameW(&ofn))
        {
            if (AddAppToContextMenu(fileName))
            {
                MessageBoxW(hMainWindow,
                            L"程序已成功添加到桌面右键菜单！\n"
                            L"程序图标也会显示在菜单中。\n"
                            L"可能需要刷新桌面或重新启动资源管理器才能看到变化。",
                            L"成功", MB_OK | MB_ICONINFORMATION);
            }
            else
            {
                MessageBoxW(hMainWindow, L"添加程序失败！请以管理员身份运行程序。", L"错误", MB_OK | MB_ICONERROR);
            }
        }
    }

    void OnRemoveButtonClick()
    {
        int selectedIndex = (int)SendMessageW(hListBox, LB_GETCURSEL, 0, 0);
        if (selectedIndex == LB_ERR)
        {
            MessageBoxW(hMainWindow, L"请先选择一个程序！", L"提示", MB_OK | MB_ICONINFORMATION);
            return;
        }

        AppEntry &app = apps[selectedIndex];
        std::wstring confirmMsg = L"确定要从桌面右键菜单中删除这个程序吗？\n\n";
        confirmMsg += L"名称: " + app.displayName + L"\n";
        confirmMsg += L"路径: " + app.path;

        if (MessageBoxW(hMainWindow, confirmMsg.c_str(), L"确认删除", MB_YESNO | MB_ICONQUESTION) == IDYES)
        {
            if (RemoveAppFromContextMenu(selectedIndex))
            {
                MessageBoxW(hMainWindow, L"程序已从桌面右键菜单中删除！", L"成功", MB_OK | MB_ICONINFORMATION);
            }
            else
            {
                MessageBoxW(hMainWindow, L"删除程序失败！", L"错误", MB_OK | MB_ICONERROR);
            }
        }
    }

    void OnRefreshButtonClick()
    {
        LoadAllContextMenuItems();
        MessageBoxW(hMainWindow, L"列表已刷新！", L"刷新", MB_OK | MB_ICONINFORMATION);
    }

    void OnShowAllCheckboxClick()
    {
        showAllItems = (SendMessage(hShowAllCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
        FilterApps();
    }

    // 静态窗口过程函数
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
    {
        RightClickManager *pThis = nullptr;

        if (uMsg == WM_NCCREATE)
        {
            CREATESTRUCTW *pCreate = reinterpret_cast<CREATESTRUCTW *>(lParam);
            pThis = reinterpret_cast<RightClickManager *>(pCreate->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pThis));
        }
        else
        {
            pThis = reinterpret_cast<RightClickManager *>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        }

        if (pThis)
        {
            return pThis->HandleMessage(hwnd, uMsg, wParam, lParam);
        }

        return DefWindowProcW(hwnd, uMsg, wParam, lParam);
    }

    // 消息处理函数
    LRESULT HandleMessage(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
    {
        switch (uMsg)
        {
        case WM_COMMAND:
            if (LOWORD(wParam) == 1002)
            { // 添加按钮
                OnAddButtonClick();
            }
            else if (LOWORD(wParam) == 1003)
            { // 删除按钮
                OnRemoveButtonClick();
            }
            else if (LOWORD(wParam) == 1004)
            { // 刷新按钮
                OnRefreshButtonClick();
            }
            else if (LOWORD(wParam) == 1005)
            { // 复选框
                OnShowAllCheckboxClick();
            }
            break;

        case WM_GETMINMAXINFO:
        {
            // 锁定窗口大小
            MINMAXINFO *mmi = (MINMAXINFO *)lParam;
            mmi->ptMinTrackSize.x = 750; // 最小宽度
            mmi->ptMinTrackSize.y = 450; // 最小高度
            mmi->ptMaxTrackSize.x = 750; // 最大宽度
            mmi->ptMaxTrackSize.y = 450; // 最大高度
        }
        break;

        case WM_DESTROY:
            PostQuitMessage(0);
            break;

        default:
            return DefWindowProcW(hwnd, uMsg, wParam, lParam);
        }
        return 0;
    }

    // 运行消息循环
    int Run()
    {
        MSG msg = {};
        while (GetMessage(&msg, NULL, 0, 0))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        return (int)msg.wParam;
    }
};

// 程序入口点
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
    RightClickManager manager;

    // 检查是否已有实例在运行
    if (manager.IsAlreadyRunning())
    {
        manager.ActivateExistingInstance();
        return 0; // 直接退出，不创建新实例
    }

    if (!manager.Initialize(hInstance))
    {
        MessageBoxW(NULL, L"程序初始化失败！", L"错误", MB_OK | MB_ICONERROR);
        return 1;
    }

    return manager.Run();
}