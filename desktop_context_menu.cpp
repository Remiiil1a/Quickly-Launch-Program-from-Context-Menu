#include <windows.h>
#include <commctrl.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <algorithm>

// 显式链接所需的库
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "shlwapi.lib")

// 应用程序结构
struct AppEntry {
    std::wstring name;
    std::wstring path;
    std::wstring displayName;
};

class RightClickManager {
private:
    std::vector<AppEntry> apps;
    HWND hMainWindow;
    HWND hListBox;
    HWND hAddButton;
    HWND hRemoveButton;

public:
    RightClickManager() : hMainWindow(NULL), hListBox(NULL), hAddButton(NULL), hRemoveButton(NULL) {}

    // 初始化程序
    bool Initialize(HINSTANCE hInstance) {
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
        int windowWidth = 600;
        int windowHeight = 400;
        int x = (screenWidth - windowWidth) / 2;
        int y = (screenHeight - windowHeight) / 2;

        // 修改窗口样式，移除可调整大小的边框
        hMainWindow = CreateWindowExW(
            0,
            L"RightClickManager",
            L"桌面右键菜单管理器",
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX, // 移除 WS_THICKFRAME 和 WS_MAXIMIZEBOX
            x, y,
            windowWidth, windowHeight,
            NULL, NULL, hInstance, this
        );

        if (!hMainWindow) return false;

        CreateControls(hInstance);
        LoadExistingApps();
        ShowWindow(hMainWindow, SW_SHOW);
        UpdateWindow(hMainWindow);

        return true;
    }

    // 创建控件
    void CreateControls(HINSTANCE hInstance) {
        // 列表控件
        hListBox = CreateWindowExW(
            WS_EX_CLIENTEDGE,
            L"LISTBOX",
            L"",
            WS_CHILD | WS_VISIBLE | LBS_NOTIFY | LBS_HASSTRINGS | WS_VSCROLL,
            10, 10, 400, 300,
            hMainWindow,
            (HMENU)1001,
            hInstance,
            NULL
        );

        // 添加按钮
        hAddButton = CreateWindowW(
            L"BUTTON",
            L"添加程序",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            420, 10, 150, 30,
            hMainWindow,
            (HMENU)1002,
            hInstance,
            NULL
        );

        // 删除按钮
        hRemoveButton = CreateWindowW(
            L"BUTTON",
            L"删除选中项",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            420, 50, 150, 30,
            hMainWindow,
            (HMENU)1003,
            hInstance,
            NULL
        );

        // 添加说明文本
        CreateWindowW(
            L"STATIC",
            L"管理桌面右键菜单中的程序快捷方式\n\n添加的程序将显示自身图标",
            WS_CHILD | WS_VISIBLE,
            420, 100, 150, 80,
            hMainWindow,
            NULL,
            hInstance,
            NULL
        );
    }

    // 加载已存在的应用
    void LoadExistingApps() {
        apps.clear();
        SendMessageW(hListBox, LB_RESETCONTENT, 0, 0);

        // 检查桌面右键菜单注册表位置
        HKEY hKey;
        if (RegOpenKeyExW(HKEY_CLASSES_ROOT, L"Directory\\Background\\shell", 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
            wchar_t subkeyName[256];
            DWORD index = 0;
            DWORD nameSize = sizeof(subkeyName) / sizeof(wchar_t);

            while (RegEnumKeyExW(hKey, index, subkeyName, &nameSize, NULL, NULL, NULL, NULL) == ERROR_SUCCESS) {
                if (wcsstr(subkeyName, L"CustomApp_") == subkeyName) {
                    AppEntry app;
                    app.name = subkeyName;
                    
                    // 获取显示名称
                    std::wstring commandPath = L"Directory\\Background\\shell\\";
                    commandPath += subkeyName;
                    commandPath += L"\\command";
                    
                    HKEY hCommandKey;
                    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, commandPath.c_str(), 0, KEY_READ, &hCommandKey) == ERROR_SUCCESS) {
                        wchar_t appPath[1024];
                        DWORD pathSize = sizeof(appPath);
                        if (RegQueryValueExW(hCommandKey, NULL, NULL, NULL, (LPBYTE)appPath, &pathSize) == ERROR_SUCCESS) {
                            app.path = appPath;
                            // 清理路径（移除可能的引号和参数）
                            CleanAppPath(app.path);
                        }
                        RegCloseKey(hCommandKey);
                    }
                    
                    // 获取显示名称
                    std::wstring displayPath = L"Directory\\Background\\shell\\";
                    displayPath += subkeyName;
                    HKEY hDisplayKey;
                    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, displayPath.c_str(), 0, KEY_READ, &hDisplayKey) == ERROR_SUCCESS) {
                        wchar_t displayName[256];
                        DWORD nameSize = sizeof(displayName);
                        if (RegQueryValueExW(hDisplayKey, NULL, NULL, NULL, (LPBYTE)displayName, &nameSize) == ERROR_SUCCESS) {
                            app.displayName = displayName;
                        } else {
                            app.displayName = L"未知程序";
                        }
                        RegCloseKey(hDisplayKey);
                    }
                    
                    apps.push_back(app);
                    
                    std::wstring listText = app.displayName + L" - " + app.path;
                    SendMessageW(hListBox, LB_ADDSTRING, 0, (LPARAM)listText.c_str());
                }
                index++;
                nameSize = sizeof(subkeyName) / sizeof(wchar_t);
            }
            RegCloseKey(hKey);
        }
    }

    // 清理应用路径
    void CleanAppPath(std::wstring& path) {
        // 移除引号
        if (path.length() >= 2 && path[0] == L'\"' && path[path.length()-1] == L'\"') {
            path = path.substr(1, path.length()-2);
        }
        // 移除参数
        size_t pos = path.find(L".exe");
        if (pos != std::wstring::npos) {
            path = path.substr(0, pos + 4);
        }
    }

    // 添加应用到桌面右键菜单
    bool AddAppToContextMenu(const std::wstring& appPath) {
        // 获取程序名称
        size_t lastSlash = appPath.find_last_of(L'\\');
        size_t lastDot = appPath.find_last_of(L'.');
        if (lastSlash == std::wstring::npos) return false;
        
        std::wstring appName = appPath.substr(lastSlash + 1);
        if (lastDot != std::wstring::npos && lastDot > lastSlash) {
            appName = appPath.substr(lastSlash + 1, lastDot - lastSlash - 1);
        }
        
        // 生成唯一的注册表键名
        std::wstring registryKey = L"CustomApp_";
        registryKey += appName;
        
        // 创建注册表项 - 这里改为桌面右键菜单的位置
        std::wstring shellKey = L"Directory\\Background\\shell\\";
        shellKey += registryKey;
        
        HKEY hKey;
        LONG result = RegCreateKeyExW(HKEY_CLASSES_ROOT, shellKey.c_str(), 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL);
        if (result == ERROR_SUCCESS) {
            // 设置显示名称
            result = RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE*)appName.c_str(), (appName.length() + 1) * sizeof(wchar_t));
            if (result != ERROR_SUCCESS) {
                RegCloseKey(hKey);
                return false;
            }
            
            // 设置图标 - 这是新增的关键部分
            // 使用程序自身的图标
            std::wstring iconValue = L"\"";
            iconValue += appPath;
            iconValue += L"\"";
            
            RegSetValueExW(hKey, L"Icon", 0, REG_SZ, (const BYTE*)iconValue.c_str(), (iconValue.length() + 1) * sizeof(wchar_t));
            
            RegCloseKey(hKey);
        } else {
            return false;
        }
        
        // 创建command子键
        std::wstring commandKey = shellKey + L"\\command";
        if (RegCreateKeyExW(HKEY_CLASSES_ROOT, commandKey.c_str(), 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
            // 桌面右键菜单不需要传递文件参数，直接启动程序即可
            std::wstring commandValue = L"\"";
            commandValue += appPath;
            commandValue += L"\"";
            
            result = RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE*)commandValue.c_str(), (commandValue.length() + 1) * sizeof(wchar_t));
            RegCloseKey(hKey);
            
            if (result == ERROR_SUCCESS) {
                // 刷新系统，使注册表更改立即生效
                SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
                
                LoadExistingApps(); // 刷新列表
                return true;
            }
        }
        
        return false;
    }

    // 从右键菜单删除应用
    bool RemoveAppFromContextMenu(int index) {
        if (index < 0 || index >= (int)apps.size()) return false;
        
        std::wstring shellKey = L"Directory\\Background\\shell\\";
        shellKey += apps[index].name;
        
        // 递归删除注册表项
        if (DeleteRegistryTree(HKEY_CLASSES_ROOT, shellKey.c_str())) {
            // 刷新系统
            SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
            
            apps.erase(apps.begin() + index);
            LoadExistingApps(); // 刷新列表
            return true;
        }
        
        return false;
    }

    // 递归删除注册表树
    bool DeleteRegistryTree(HKEY hParentKey, const wchar_t* subkey) {
        HKEY hKey;
        if (RegOpenKeyExW(hParentKey, subkey, 0, KEY_READ | KEY_WRITE, &hKey) != ERROR_SUCCESS) {
            return false;
        }
        
        // 枚举并删除所有子键
        wchar_t childKeyName[256];
        DWORD childKeySize = sizeof(childKeyName) / sizeof(wchar_t);
        
        while (RegEnumKeyExW(hKey, 0, childKeyName, &childKeySize, NULL, NULL, NULL, NULL) == ERROR_SUCCESS) {
            std::wstring fullChildKey = subkey;
            fullChildKey += L"\\";
            fullChildKey += childKeyName;
            
            if (!DeleteRegistryTree(hParentKey, fullChildKey.c_str())) {
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
    void OnAddButtonClick() {
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

        if (GetOpenFileNameW(&ofn)) {
            if (AddAppToContextMenu(fileName)) {
                MessageBoxW(hMainWindow, 
                    L"程序已成功添加到桌面右键菜单！\n"
                    L"程序图标也会显示在菜单中。\n"
                    L"可能需要刷新桌面或重新启动资源管理器才能看到变化。", 
                    L"成功", MB_OK | MB_ICONINFORMATION);
            } else {
                MessageBoxW(hMainWindow, L"添加程序失败！请以管理员身份运行程序。", L"错误", MB_OK | MB_ICONERROR);
            }
        }
    }

    void OnRemoveButtonClick() {
        int selectedIndex = (int)SendMessageW(hListBox, LB_GETCURSEL, 0, 0);
        if (selectedIndex == LB_ERR) {
            MessageBoxW(hMainWindow, L"请先选择一个程序！", L"提示", MB_OK | MB_ICONINFORMATION);
            return;
        }

        if (MessageBoxW(hMainWindow, L"确定要从桌面右键菜单中删除这个程序吗？", L"确认删除", MB_YESNO | MB_ICONQUESTION) == IDYES) {
            if (RemoveAppFromContextMenu(selectedIndex)) {
                MessageBoxW(hMainWindow, L"程序已从桌面右键菜单中删除！", L"成功", MB_OK | MB_ICONINFORMATION);
            } else {
                MessageBoxW(hMainWindow, L"删除程序失败！", L"错误", MB_OK | MB_ICONERROR);
            }
        }
    }

    // 静态窗口过程函数
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
        RightClickManager* pThis = nullptr;

        if (uMsg == WM_NCCREATE) {
            CREATESTRUCTW* pCreate = reinterpret_cast<CREATESTRUCTW*>(lParam);
            pThis = reinterpret_cast<RightClickManager*>(pCreate->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pThis));
        } else {
            pThis = reinterpret_cast<RightClickManager*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        }

        if (pThis) {
            return pThis->HandleMessage(hwnd, uMsg, wParam, lParam);
        }

        return DefWindowProcW(hwnd, uMsg, wParam, lParam);
    }

    // 消息处理函数
    LRESULT HandleMessage(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
        switch (uMsg) {
        case WM_COMMAND:
            if (LOWORD(wParam) == 1002) { // 添加按钮
                OnAddButtonClick();
            } else if (LOWORD(wParam) == 1003) { // 删除按钮
                OnRemoveButtonClick();
            }
            break;

        case WM_GETMINMAXINFO:
            {
                // 锁定窗口大小
                MINMAXINFO* mmi = (MINMAXINFO*)lParam;
                mmi->ptMinTrackSize.x = 600;  // 最小宽度
                mmi->ptMinTrackSize.y = 400;  // 最小高度
                mmi->ptMaxTrackSize.x = 600;  // 最大宽度
                mmi->ptMaxTrackSize.y = 400;  // 最大高度
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
    int Run() {
        MSG msg = {};
        while (GetMessage(&msg, NULL, 0, 0)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        return (int)msg.wParam;
    }
};

// 程序入口点
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    RightClickManager manager;
    
    if (!manager.Initialize(hInstance)) {
        MessageBoxW(NULL, L"程序初始化失败！", L"错误", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    return manager.Run();
}