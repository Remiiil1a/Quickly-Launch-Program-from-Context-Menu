#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0A00 // 将其定义为Windows 10
#endif

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <algorithm>
#include <map>
#include <shlwapi.h>
#include <shellscalingapi.h>

#define IDI_MAIN_ICON 101
#define IDI_SMALL_ICON 102

#ifndef GET_X_LPARAM
#define GET_X_LPARAM(lParam) ((int)(short)LOWORD(lParam))
#endif

#ifndef GET_Y_LPARAM
#define GET_Y_LPARAM(lParam) ((int)(short)HIWORD(lParam))
#endif

// 显式链接所需的库
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "Shcore.lib")

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
    HWND hEditBox; // 编辑框句柄
    HANDLE hMutex;
    bool showAllItems;    // 是否显示所有项
    bool isEditing;       // 是否正在编辑状态
    HFONT hModernFont;    // 添加字体句柄
    int editingIndex;     // 正在编辑的项索引
    WNDPROC oldEditProc;  // 保存原来的编辑框过程
    HMENU hContextMenu;   // 右键菜单句柄
    int contextMenuIndex; // 右键菜单对应的项索引

    static LRESULT CALLBACK EditBoxProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
    {
        RightClickManager *pThis = (RightClickManager *)GetWindowLongPtr(hwnd, GWLP_USERDATA);

        if (pThis && pThis->isEditing)
        {
            switch (uMsg)
            {
            case WM_KEYDOWN:
                if (wParam == VK_RETURN)
                {
                    pThis->FinishEditing(true);
                    return 0;
                }
                else if (wParam == VK_ESCAPE)
                {
                    pThis->CancelEditing();
                    return 0;
                }
                break;

            case WM_KILLFOCUS:
                // 失去焦点时自动保存
                pThis->FinishEditing(true);
                return 0;
            }
        }

        // 调用原来的窗口过程
        if (pThis && pThis->oldEditProc)
        {
            return CallWindowProc(pThis->oldEditProc, hwnd, uMsg, wParam, lParam);
        }

        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }

    // 按字母顺序排序应用
    void SortAppsAlphabetically()
    {
        std::sort(allApps.begin(), allApps.end(), [](const AppEntry &a, const AppEntry &b)
                  {
            // 不区分大小写比较
            return _wcsicmp(a.displayName.c_str(), b.displayName.c_str()) < 0; });
    }

    // 从注册表刷新单个项的显示名称
    void RefreshSingleItemFromRegistry(const std::wstring &itemName)
    {
        // 在 allApps 中查找对应的项
        for (auto &app : allApps)
        {
            if (app.name == itemName)
            {
                // 从注册表重新读取显示名称
                std::wstring displayPath = L"Directory\\Background\\shell\\";
                displayPath += itemName;

                HKEY hDisplayKey;
                if (RegOpenKeyExW(HKEY_CLASSES_ROOT, displayPath.c_str(), 0, KEY_READ, &hDisplayKey) == ERROR_SUCCESS)
                {
                    wchar_t displayName[256];
                    DWORD nameSize = sizeof(displayName);
                    if (RegQueryValueExW(hDisplayKey, NULL, NULL, NULL, (LPBYTE)displayName, &nameSize) == ERROR_SUCCESS)
                    {
                        // 更新内存中的显示名称
                        app.displayName = displayName;
                    }
                    RegCloseKey(hDisplayKey);
                }
                break;
            }
        }

        // 同时更新 apps 中的对应项
        for (auto &app : apps)
        {
            if (app.name == itemName)
            {
                // 从注册表重新读取显示名称
                std::wstring displayPath = L"Directory\\Background\\shell\\";
                displayPath += itemName;

                HKEY hDisplayKey;
                if (RegOpenKeyExW(HKEY_CLASSES_ROOT, displayPath.c_str(), 0, KEY_READ, &hDisplayKey) == ERROR_SUCCESS)
                {
                    wchar_t displayName[256];
                    DWORD nameSize = sizeof(displayName);
                    if (RegQueryValueExW(hDisplayKey, NULL, NULL, NULL, (LPBYTE)displayName, &nameSize) == ERROR_SUCCESS)
                    {
                        // 更新内存中的显示名称
                        app.displayName = displayName;
                    }
                    RegCloseKey(hDisplayKey);
                }
                break;
            }
        }

        // 刷新列表显示
        FilterApps();
    }

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

private:
    // 生成排序后的注册表项名称
    std::wstring GenerateSortedRegistryKey(const std::wstring &displayName)
    {
        // 获取所有现有的自定义应用
        std::vector<AppEntry> customApps;
        for (const auto &app : allApps)
        {
            if (app.isCustom)
            {
                customApps.push_back(app);
            }
        }

        // 添加新应用到临时列表
        AppEntry newApp;
        newApp.displayName = displayName;
        customApps.push_back(newApp);

        // 按显示名称排序
        std::sort(customApps.begin(), customApps.end(), [](const AppEntry &a, const AppEntry &b)
                  { return _wcsicmp(a.displayName.c_str(), b.displayName.c_str()) < 0; });

        // 找到新应用的排序位置
        int position = 1;
        for (const auto &app : customApps)
        {
            if (app.displayName == displayName)
            {
                // 生成带排序数字的注册表项名称
                wchar_t keyName[256];
                swprintf(keyName, 256, L"%02d_CustomApp_%s", position, displayName.c_str());
                return keyName;
            }
            position++;
        }

        // 默认回退
        return L"99_CustomApp_" + displayName;
    }

    // 重新排序所有注册表项 - 改进版本
    void ReorderRegistryItems()
    {
        // 只重新排序自定义应用
        std::vector<AppEntry> customApps;
        for (const auto &app : allApps)
        {
            if (app.isCustom)
            {
                customApps.push_back(app);
            }
        }

        // 按显示名称排序
        std::sort(customApps.begin(), customApps.end(), [](const AppEntry &a, const AppEntry &b)
                  { return _wcsicmp(a.displayName.c_str(), b.displayName.c_str()) < 0; });

        // 重新生成所有自定义项的注册表键名
        int position = 1;
        std::vector<std::pair<std::wstring, std::wstring>> keyMappings; // <oldKey, newKey>

        for (const auto &app : customApps)
        {
            // 生成新的注册表项名称
            wchar_t newKeyName[256];
            swprintf(newKeyName, 256, L"%02d_CustomApp_%s", position, app.displayName.c_str());

            // 只有当名称确实改变时才需要处理
            if (app.name != newKeyName)
            {
                keyMappings.push_back({app.name, newKeyName});
            }
            position++;
        }

        // 如果没有需要重命名的项，直接返回
        if (keyMappings.empty())
        {
            return;
        }

        // 为每个需要重命名的项创建新键并复制数据
        for (const auto &mapping : keyMappings)
        {
            const std::wstring &oldKeyName = mapping.first;
            const std::wstring &newKeyName = mapping.second;

            // 找到对应的应用数据
            AppEntry *app = nullptr;
            for (auto &a : allApps)
            {
                if (a.name == oldKeyName && a.isCustom)
                {
                    app = &a;
                    break;
                }
            }

            if (!app)
                continue;

            // 创建新的注册表项
            std::wstring newShellKey = L"Directory\\Background\\shell\\";
            newShellKey += newKeyName;

            HKEY hNewKey;
            if (RegCreateKeyExW(HKEY_CLASSES_ROOT, newShellKey.c_str(), 0, NULL, 0, KEY_WRITE, NULL, &hNewKey, NULL) == ERROR_SUCCESS)
            {
                // 复制显示名称 - 直接从内存中读取，确保一致性
                RegSetValueExW(hNewKey, NULL, 0, REG_SZ,
                               (const BYTE *)app->displayName.c_str(),
                               (app->displayName.length() + 1) * sizeof(wchar_t));

                // 复制图标
                if (!app->icon.empty())
                {
                    RegSetValueExW(hNewKey, L"Icon", 0, REG_SZ,
                                   (const BYTE *)app->icon.c_str(),
                                   (app->icon.length() + 1) * sizeof(wchar_t));
                }

                RegCloseKey(hNewKey);

                // 创建command子键
                std::wstring newCommandKey = newShellKey + L"\\command";
                if (RegCreateKeyExW(HKEY_CLASSES_ROOT, newCommandKey.c_str(), 0, NULL, 0, KEY_WRITE, NULL, &hNewKey, NULL) == ERROR_SUCCESS)
                {
                    std::wstring commandValue = L"\"";
                    commandValue += app->path;
                    commandValue += L"\"";

                    RegSetValueExW(hNewKey, NULL, 0, REG_SZ,
                                   (const BYTE *)commandValue.c_str(),
                                   (commandValue.length() + 1) * sizeof(wchar_t));
                    RegCloseKey(hNewKey);

                    // 更新内存中的名称
                    app->name = newKeyName;
                }
                else
                {
                    // 创建command失败，删除shell键
                    RegDeleteKeyW(HKEY_CLASSES_ROOT, newShellKey.c_str());
                    continue;
                }
            }
            else
            {
                continue;
            }
        }

        // 删除所有旧的注册表项（只删除成功创建了新键的旧键）
        for (const auto &mapping : keyMappings)
        {
            std::wstring oldShellKey = L"Directory\\Background\\shell\\";
            oldShellKey += mapping.first;

            // 确保新键存在后再删除旧键
            std::wstring newShellKey = L"Directory\\Background\\shell\\";
            newShellKey += mapping.second;

            HKEY hTestKey;
            if (RegOpenKeyExW(HKEY_CLASSES_ROOT, newShellKey.c_str(), 0, KEY_READ, &hTestKey) == ERROR_SUCCESS)
            {
                RegCloseKey(hTestKey);
                // 新键存在，可以安全删除旧键
                DeleteRegistryTree(HKEY_CLASSES_ROOT, oldShellKey.c_str());
            }
        }

        // 刷新系统
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

        // 重新加载所有菜单项 - 从注册表重新读取确保数据一致
        LoadAllContextMenuItems();
    }

    // 强制从注册表重新加载所有菜单项
    void ForceReloadFromRegistry()
    {
        // 清空现有数据
        allApps.clear();
        apps.clear();

        // 重置列表框
        SendMessageW(hListBox, LB_RESETCONTENT, 0, 0);

        // 直接从注册表重新读取
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
                    // 修改自定义应用识别逻辑，支持新的排序命名格式
                    app.isCustom = (wcsstr(subkeyName, L"CustomApp_") != nullptr);

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
                            app.displayName = subkeyName;
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

        // 重新排序并过滤应用列表
        SortAppsAlphabetically();
        FilterApps();
    }

public:
    RightClickManager() : hMainWindow(NULL), hListBox(NULL), hAddButton(NULL),
                          hRemoveButton(NULL), hRefreshButton(NULL), hShowAllCheckbox(NULL),
                          hEditBox(NULL), hMutex(NULL), showAllItems(false), isEditing(false),
                          hModernFont(NULL), editingIndex(-1), oldEditProc(NULL),
                          hContextMenu(NULL), contextMenuIndex(-1) {}

    ~RightClickManager()
    {
        if (hMutex)
        {
            CloseHandle(hMutex);
            hMutex = NULL;
        }
        if (hModernFont)
        {
            DeleteObject(hModernFont);
            hModernFont = NULL;
        }
        if (hContextMenu)
        {
            DestroyMenu(hContextMenu);
            hContextMenu = NULL;
        }
    }

    // 创建上下文菜单
    void CreateContextMenu()
    {
        hContextMenu = CreatePopupMenu();
        if (hContextMenu)
        {
            AppendMenuW(hContextMenu, MF_STRING, 1101, L"📁 在注册表中打开");
            AppendMenuW(hContextMenu, MF_SEPARATOR, 0, NULL);
            AppendMenuW(hContextMenu, MF_STRING, 1102, L"🔄 刷新此项");
        }
    }

    // 显示上下文菜单
    void ShowContextMenu(int x, int y, int index)
    {
        if (index < 0 || index >= (int)apps.size())
            return;

        contextMenuIndex = index;

        if (!hContextMenu)
        {
            CreateContextMenu();
        }

        if (hContextMenu)
        {
            // 获取列表项的矩形位置
            RECT itemRect;
            if (SendMessageW(hListBox, LB_GETITEMRECT, index, (LPARAM)&itemRect) != LB_ERR)
            {
                // 计算菜单应该显示的位置（在项的下方）
                POINT pt;
                pt.x = itemRect.left + 10; // 稍微偏移避免覆盖文本
                pt.y = itemRect.bottom;

                // 转换为屏幕坐标
                ClientToScreen(hListBox, &pt);

                // 显示右键菜单
                TrackPopupMenuEx(hContextMenu,
                                 TPM_RIGHTBUTTON | TPM_LEFTALIGN | TPM_TOPALIGN,
                                 pt.x, pt.y, hMainWindow, NULL);
            }
            else
            {
                // 备用方法：使用鼠标位置
                POINT pt = {x, y};
                ClientToScreen(hListBox, &pt);
                TrackPopupMenuEx(hContextMenu,
                                 TPM_RIGHTBUTTON | TPM_LEFTALIGN | TPM_TOPALIGN,
                                 pt.x, pt.y, hMainWindow, NULL);
            }
        }
    }

    // 打开注册表位置
    void OpenRegistryLocation(int index)
    {
        if (index < 0 || index >= (int)apps.size())
            return;

        AppEntry &app = apps[index];

        // 构建注册表路径
        std::wstring regPath = L"计算机\\HKEY_CLASSES_ROOT\\Directory\\Background\\shell\\" + app.name;

        // 尝试使用 ShellExecute 打开注册表编辑器
        SHELLEXECUTEINFOW sei = {sizeof(sei)};
        sei.lpVerb = L"open";
        sei.lpFile = L"regedit.exe";
        sei.lpParameters = L""; // 不使用静默模式，以便用户看到界面
        sei.nShow = SW_SHOW;

        if (ShellExecuteExW(&sei))
        {
            // 给注册表编辑器一些时间启动
            Sleep(1000);

            // 尝试发送按键到注册表编辑器来导航到我们的路径
            HWND hRegEdit = FindWindowW(L"RegEdit_RegEdit", NULL);
            if (hRegEdit)
            {
                // 激活窗口
                SetForegroundWindow(hRegEdit);
                Sleep(100);

                // 发送 Ctrl+F 打开查找
                keybd_event(VK_CONTROL, 0, 0, 0);
                keybd_event('F', 0, 0, 0);
                keybd_event('F', 0, KEYEVENTF_KEYUP, 0);
                keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);

                // 等待查找对话框打开
                Sleep(500);

                // 在查找框中输入我们的路径
                // 首先确保输入法是英文状态
                HWND hForeground = GetForegroundWindow();

                // 使用更可靠的方式输入路径
                for (wchar_t c : regPath)
                {
                    // 对于反斜杠，使用虚拟键码而不是字符，避免输入法干扰
                    if (c == L'\\')
                    {
                        // 直接发送反斜杠的虚拟键码
                        keybd_event(VK_OEM_5, 0, 0, 0);
                        keybd_event(VK_OEM_5, 0, KEYEVENTF_KEYUP, 0);
                    }
                    else if (c == L'：') // 中文冒号转英文冒号
                    {
                        keybd_event(VK_SHIFT, 0, 0, 0);
                        keybd_event(VK_OEM_1, 0, 0, 0);
                        keybd_event(VK_OEM_1, 0, KEYEVENTF_KEYUP, 0);
                        keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
                    }
                    else
                    {
                        // 对于其他字符，使用更可靠的输入方式
                        SHORT vk = VkKeyScanW(c);
                        if (vk != -1)
                        {
                            BYTE vkCode = LOBYTE(vk);
                            BYTE shiftState = HIBYTE(vk);

                            if (shiftState & 1) // Shift 键
                            {
                                keybd_event(VK_SHIFT, 0, 0, 0);
                            }
                            if (shiftState & 2) // Ctrl 键
                            {
                                keybd_event(VK_CONTROL, 0, 0, 0);
                            }
                            if (shiftState & 4) // Alt 键
                            {
                                keybd_event(VK_MENU, 0, 0, 0);
                            }

                            keybd_event(vkCode, 0, 0, 0);
                            keybd_event(vkCode, 0, KEYEVENTF_KEYUP, 0);

                            if (shiftState & 4)
                            {
                                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
                            }
                            if (shiftState & 2)
                            {
                                keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
                            }
                            if (shiftState & 1)
                            {
                                keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
                            }
                        }
                    }
                    Sleep(10); // 小延迟确保字符正确输入
                }

                // 发送回车开始查找
                Sleep(100);
                keybd_event(VK_RETURN, 0, 0, 0);
                keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0);

                std::wstring message = L"注册表编辑器已打开并尝试定位到指定位置。\n"
                                       L"如果未自动定位，请手动导航到：\n" +
                                       regPath;

                MessageBoxW(hMainWindow,
                            message.c_str(),
                            L"打开注册表位置",
                            MB_OK | MB_ICONINFORMATION);
            }
            else
            {
                std::wstring message = L"注册表编辑器已打开，但无法自动定位。\n"
                                       L"请手动导航到：\n" +
                                       regPath;

                MessageBoxW(hMainWindow,
                            message.c_str(),
                            L"打开注册表位置",
                            MB_OK | MB_ICONINFORMATION);
            }
        }
        else
        {
            // 备用方法：显示路径让用户手动导航
            std::wstring message = L"请手动在注册表编辑器中导航到以下路径：\n\n" + regPath;

            MessageBoxW(hMainWindow,
                        message.c_str(),
                        L"注册表位置",
                        MB_OK | MB_ICONINFORMATION);
        }
    }

    // 处理列表框双击事件
    void OnListBoxDoubleClick()
    {
        if (isEditing)
            return; // 如果正在编辑，忽略双击

        int selectedIndex = (int)SendMessageW(hListBox, LB_GETCURSEL, 0, 0);
        if (selectedIndex == LB_ERR)
            return;

        // 只允许重命名本程序创建的项
        if (selectedIndex >= 0 && selectedIndex < (int)apps.size() && apps[selectedIndex].isCustom)
        {
            StartEditing(selectedIndex);
        }
        else
        {
            MessageBoxW(hMainWindow,
                        L"只能重命名本程序创建的项目（✅ 标记的项）",
                        L"提示",
                        MB_OK | MB_ICONINFORMATION);
        }
    }

    // 开始编辑
    void StartEditing(int index)
    {
        editingIndex = index;
        isEditing = true;

        // 获取列表项的位置
        RECT itemRect;
        SendMessageW(hListBox, LB_GETITEMRECT, index, (LPARAM)&itemRect);

        // 调整编辑框位置，避开图标区域
        itemRect.left += 30; // 为图标留出空间

        // 创建编辑框
        hEditBox = CreateWindowW(
            L"EDIT",
            apps[index].displayName.c_str(),
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_BORDER,
            itemRect.left, itemRect.top,
            itemRect.right - itemRect.left - 5, itemRect.bottom - itemRect.top,
            hListBox,
            NULL,
            (HINSTANCE)GetWindowLongPtr(hListBox, GWLP_HINSTANCE),
            NULL);

        if (hEditBox && hModernFont)
        {
            SendMessage(hEditBox, WM_SETFONT, (WPARAM)hModernFont, TRUE);
            SetFocus(hEditBox);
            SendMessage(hEditBox, EM_SETSEL, 0, -1); // 全选文本

            // 子类化编辑框以捕获键盘消息
            SetWindowLongPtr(hEditBox, GWLP_USERDATA, (LONG_PTR)this);
            oldEditProc = (WNDPROC)SetWindowLongPtr(hEditBox, GWLP_WNDPROC, (LONG_PTR)EditBoxProc);
        }
    }

    // 完成编辑 - 改进版本
    void FinishEditing(bool saveChanges)
    {
        if (!isEditing)
            return;

        if (saveChanges && hEditBox)
        {
            // 获取编辑框中的新名称
            wchar_t newName[256];
            GetWindowTextW(hEditBox, newName, 256);

            if (wcslen(newName) > 0)
            {
                AppEntry &app = apps[editingIndex];
                std::wstring oldDisplayName = app.displayName;

                // 检查名称是否真的改变了
                if (lstrcmpiW(oldDisplayName.c_str(), newName) != 0)
                {
                    // 更新注册表中的显示名称
                    std::wstring shellKey = L"Directory\\Background\\shell\\";
                    shellKey += app.name;

                    HKEY hKey;
                    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, shellKey.c_str(), 0, KEY_WRITE, &hKey) == ERROR_SUCCESS)
                    {
                        RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE *)newName, (wcslen(newName) + 1) * sizeof(wchar_t));
                        RegCloseKey(hKey);

                        // 刷新系统
                        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

                        // 直接从注册表重新读取该项的显示名称，确保数据同步
                        RefreshSingleItemFromRegistry(app.name);

                        // 重新排序注册表项
                        ReorderRegistryItems();

                        MessageBoxW(hMainWindow, L"重命名成功！", L"成功", MB_OK | MB_ICONINFORMATION);
                    }
                    else
                    {
                        MessageBoxW(hMainWindow, L"重命名失败！请以管理员身份运行。", L"错误", MB_OK | MB_ICONERROR);
                    }
                }
                else
                {
                    // 名称没有改变，不需要做任何操作
                    MessageBoxW(hMainWindow, L"名称没有改变。", L"提示", MB_OK | MB_ICONINFORMATION);
                }
            }
        }

        // 清理编辑状态
        if (hEditBox)
        {
            if (oldEditProc)
            {
                SetWindowLongPtr(hEditBox, GWLP_WNDPROC, (LONG_PTR)oldEditProc);
                oldEditProc = NULL;
            }
            DestroyWindow(hEditBox);
            hEditBox = NULL;
        }

        editingIndex = -1;
        isEditing = false;
        SetFocus(hListBox);
    }

    // 取消编辑
    void CancelEditing()
    {
        FinishEditing(false);
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
        // 创建现代字体
        NONCLIENTMETRICSW ncm = {sizeof(ncm)};
        SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
        hModernFont = CreateFontIndirectW(&ncm.lfMessageFont);

        // 初始化通用控件
        INITCOMMONCONTROLSEX icex;
        icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
        icex.dwICC = ICC_LISTVIEW_CLASSES | ICC_TREEVIEW_CLASSES | ICC_BAR_CLASSES;
        InitCommonControlsEx(&icex);

        // 获取DPI缩放因子
        HDC hdc = GetDC(NULL);
        int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
        ReleaseDC(NULL, hdc);
        float scale = dpiX / 96.0f;

        // 创建主窗口
        WNDCLASSEXW wc = {};
        wc.cbSize = sizeof(WNDCLASSEXW);
        wc.style = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc = WindowProc;
        wc.hInstance = hInstance;
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"RightClickManager";
        wc.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MAIN_ICON));
        wc.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL_ICON));

        RegisterClassExW(&wc);

        // 根据DPI缩放计算窗口大小
        int windowWidth = (int)(750 * scale);  // 增加基础宽度
        int windowHeight = (int)(500 * scale); // 增加基础高度

        // 计算窗口位置使其居中
        int screenWidth = GetSystemMetrics(SM_CXSCREEN);
        int screenHeight = GetSystemMetrics(SM_CYSCREEN);
        int x = (screenWidth - windowWidth) / 2;
        int y = (screenHeight - windowHeight) / 2;

        // 修改窗口样式，移除可调整大小的边框
        hMainWindow = CreateWindowExW(
            WS_EX_APPWINDOW,
            L"RightClickManager",
            L"桌面右键菜单管理器",
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
            x, y,
            windowWidth, windowHeight,
            NULL, NULL, hInstance, this);

        if (!hMainWindow)
            return false;

        // 在主窗口创建后应用字体
        if (hModernFont)
        {
            SendMessage(hMainWindow, WM_SETFONT, (WPARAM)hModernFont, TRUE);
        }

        CreateControls(hInstance);
        LoadAllContextMenuItems();
        ShowWindow(hMainWindow, SW_SHOW);
        UpdateWindow(hMainWindow);

        return true;
    }

    // 创建控件
    void CreateControls(HINSTANCE hInstance)
    {
        // 获取DPI缩放因子
        HDC hdc = GetDC(hMainWindow);
        int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
        ReleaseDC(hMainWindow, hdc);
        float scale = dpiX / 96.0f;

        // 根据DPI缩放调整尺寸
        int listBoxWidth = (int)(580 * scale);
        int listBoxHeight = (int)(380 * scale);
        int buttonWidth = (int)(140 * scale);
        int buttonHeight = (int)(30 * scale);
        int margin = (int)(10 * scale);
        int rightPanelX = (int)(600 * scale);
        int helpTextWidth = (int)(140 * scale);
        int helpTextHeight = (int)(200 * scale);

        // 列表控件 - 确保包含垂直和水平滚动条
        hListBox = CreateWindowExW(
            WS_EX_CLIENTEDGE,
            L"LISTBOX",
            L"",
            WS_CHILD | WS_VISIBLE | LBS_NOTIFY | LBS_HASSTRINGS |
                WS_VSCROLL | WS_HSCROLL | LBS_NOINTEGRALHEIGHT | LBS_DISABLENOSCROLL, // 添加 LBS_DISABLENOSCROLL 确保滚动条始终可用
            margin, margin,
            listBoxWidth, listBoxHeight,
            hMainWindow,
            (HMENU)1001,
            hInstance,
            NULL);

        // 设置列表项高度，使文本更容易阅读
        int itemHeight = (int)(24 * scale); // 稍微增加项高度以改善可读性
        SendMessage(hListBox, LB_SETITEMHEIGHT, 0, itemHeight);

        // 添加按钮
        hAddButton = CreateWindowW(
            L"BUTTON",
            L"📁 添加程序",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            rightPanelX, margin,
            buttonWidth, buttonHeight,
            hMainWindow,
            (HMENU)1002,
            hInstance,
            NULL);

        // 删除按钮
        hRemoveButton = CreateWindowW(
            L"BUTTON",
            L"🗑️ 删除选中",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            rightPanelX, margin + buttonHeight + margin / 2,
            buttonWidth, buttonHeight,
            hMainWindow,
            (HMENU)1003,
            hInstance,
            NULL);

        // 刷新按钮
        hRefreshButton = CreateWindowW(
            L"BUTTON",
            L"🔄 刷新列表",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            rightPanelX, margin + (buttonHeight + margin / 2) * 2,
            buttonWidth, buttonHeight,
            hMainWindow,
            (HMENU)1004,
            hInstance,
            NULL);

        // 创建复选框
        hShowAllCheckbox = CreateWindowW(
            L"BUTTON",
            L"显示所有项目",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
            rightPanelX, margin + (buttonHeight + margin / 2) * 3,
            buttonWidth, buttonHeight,
            hMainWindow,
            (HMENU)1005,
            hInstance,
            NULL);

        // 帮助文本 - 简化文本以适应空间
        HWND hHelpText = CreateWindowW(
            L"STATIC",
            L"💡 桌面右键菜单管理\n\n"
            L"✅ 本程序创建的项目\n"
            L"📌 其他程序创建的项目\n\n"
            L"🖱️ 操作提示:\n"
            L"• 双击 ✅ 项可以重命名\n"
            L"• 右键项打开功能菜单\n"
            L"• 勾选复选框显示所有项目",
            WS_CHILD | WS_VISIBLE,
            rightPanelX, margin + (buttonHeight + margin / 2) * 4 + 10,
            helpTextWidth, helpTextHeight,
            hMainWindow,
            NULL,
            hInstance,
            NULL);

        // 应用现代字体到所有控件
        HWND hControls[] = {hListBox, hAddButton, hRemoveButton, hRefreshButton, hShowAllCheckbox, hHelpText};
        for (HWND hControl : hControls)
        {
            if (hControl && hModernFont)
            {
                SendMessage(hControl, WM_SETFONT, (WPARAM)hModernFont, TRUE);
            }
        }

        SendMessage(hShowAllCheckbox, BM_SETCHECK, showAllItems ? BST_CHECKED : BST_UNCHECKED, 0);
    }

    std::wstring GetDisplayText(const AppEntry &app, int maxDisplayLength = 100)
    {
        std::wstring baseText = app.isCustom ? L"✅ " : L"📌 ";
        baseText += app.displayName + L" - " + app.path;

        // 如果文本过长，进行适当截断（但这只是显示，完整内容仍可通过滚动查看）
        if (baseText.length() > maxDisplayLength)
        {
            // 保留开头和结尾的重要信息
            std::wstring shortened = baseText.substr(0, maxDisplayLength - 10) + L"..." +
                                     baseText.substr(baseText.length() - 7);
            return shortened;
        }

        return baseText;
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

        // 如果正在编辑，先取消编辑
        if (isEditing)
        {
            CancelEditing();
        }

        // 清空现有数据
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
                    app.isCustom = (wcsstr(subkeyName, L"CustomApp_") != nullptr);

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

        // 按显示名称字母顺序排序
        SortAppsAlphabetically();

        // 根据显示设置过滤应用列表
        FilterApps();
    }

    // 根据显示设置过滤应用列表
    // 根据显示设置过滤应用列表
    void FilterApps()
    {
        // 如果正在编辑，取消编辑状态
        if (isEditing)
        {
            CancelEditing();
        }

        apps.clear();
        SendMessageW(hListBox, LB_RESETCONTENT, 0, 0);

        for (const auto &app : allApps)
        {
            if (showAllItems || app.isCustom)
            {
                apps.push_back(app);

                // 使用优化后的显示文本
                std::wstring listText = GetDisplayText(app);
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
                HFONT hOldFont = (HFONT)SelectObject(hdc, hModernFont); // 确保使用正确的字体

                for (const auto &app : apps)
                {
                    std::wstring listText = app.isCustom ? L"✅ " : L"📌 ";
                    listText += app.displayName + L" - " + app.path;

                    SIZE size;
                    if (GetTextExtentPoint32W(hdc, listText.c_str(), (int)listText.length(), &size))
                    {
                        if (size.cx > maxWidth)
                        {
                            maxWidth = size.cx;
                        }
                    }
                }

                SelectObject(hdc, hOldFont);
                ReleaseDC(hListBox, hdc);

                // 设置水平滚动范围，增加足够的边距
                // 计算DPI缩放
                HDC hdcWindow = GetDC(hMainWindow);
                int dpiX = GetDeviceCaps(hdcWindow, LOGPIXELSX);
                ReleaseDC(hMainWindow, hdcWindow);
                float scale = dpiX / 96.0f;

                // 增加更多边距确保长文本完全可见
                int horizontalExtent = maxWidth + (int)(100 * scale);
                SendMessage(hListBox, LB_SETHORIZONTALEXTENT, horizontalExtent, 0);
            }
        }

        // 确保垂直滚动条在项目过多时正确显示
        // 获取列表框的客户区高度
        RECT listRect;
        GetClientRect(hListBox, &listRect);
        int listHeight = listRect.bottom - listRect.top;

        // 获取列表项高度
        int itemHeight = (int)SendMessage(hListBox, LB_GETITEMHEIGHT, 0, 0);
        if (itemHeight <= 0)
            itemHeight = 20; // 默认值

        // 计算可视区域能容纳的项目数量
        int visibleItems = listHeight / itemHeight;

        // 如果项目数量超过可视区域，垂直滚动条会自动显示
        // 我们可以通过选中第一个项目来给用户视觉反馈
        if (!apps.empty())
        {
            SendMessageW(hListBox, LB_SETCURSEL, 0, 0);
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

        // 生成带排序数字的注册表键名
        std::wstring registryKey = GenerateSortedRegistryKey(appName);

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
                // 添加错误信息
                wchar_t errorMsg[256];
                swprintf(errorMsg, 256, L"设置显示名称失败！错误代码: %d", result);
                MessageBoxW(hMainWindow, errorMsg, L"错误", MB_OK | MB_ICONERROR);
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
            // 添加错误信息
            wchar_t errorMsg[256];
            swprintf(errorMsg, 256, L"创建注册表项失败！错误代码: %d", result);
            MessageBoxW(hMainWindow, errorMsg, L"错误", MB_OK | MB_ICONERROR);
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

                // 重新排序注册表项 - 这会调用 LoadAllContextMenuItems()
                ReorderRegistryItems();
                return true;
            }
            else
            {
                // 添加错误信息
                wchar_t errorMsg[256];
                swprintf(errorMsg, 256, L"设置命令失败！错误代码: %d", result);
                MessageBoxW(hMainWindow, errorMsg, L"错误", MB_OK | MB_ICONERROR);
            }
        }
        else
        {
            // 添加错误信息
            wchar_t errorMsg[256];
            swprintf(errorMsg, 256, L"创建命令子键失败！错误代码: %d", GetLastError());
            MessageBoxW(hMainWindow, errorMsg, L"错误", MB_OK | MB_ICONERROR);
        }

        return false;
    }

    // 从右键菜单删除应用
    bool RemoveAppFromContextMenu(int index)
    {
        if (index < 0 || index >= (int)apps.size())
            return false;

        AppEntry &app = apps[index];

        // 保存要删除的项信息
        std::wstring deleteName = app.name;
        std::wstring deleteDisplayName = app.displayName;
        std::wstring deletePath = app.path;

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

        // 先尝试正常删除
        bool deleteSuccess = DeleteRegistryTree(HKEY_CLASSES_ROOT, shellKey.c_str());

        if (!deleteSuccess)
        {
            // 如果正常删除失败，显示详细错误信息
            DWORD errorCode = GetLastError();
            wchar_t errorMsg[512];
            swprintf(errorMsg, 512,
                     L"删除失败！错误代码: %d\n\n"
                     L"可能的原因：\n"
                     L"• 注册表项被其他进程占用\n"
                     L"• 权限不足\n"
                     L"• 注册表项不存在\n\n"
                     L"请尝试以管理员身份运行程序，或重启后重试。",
                     errorCode);

            MessageBoxW(hMainWindow, errorMsg, L"删除失败", MB_OK | MB_ICONERROR);
            return false;
        }

        // 刷新系统
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

        // 直接从注册表重新加载
        ForceReloadFromRegistry();

        // 显示成功消息
        MessageBoxW(hMainWindow,
                    L"程序已从桌面右键菜单中删除！\n"
                    L"如果菜单项仍然显示，请尝试刷新桌面(F5)或重启资源管理器。",
                    L"删除成功", MB_OK | MB_ICONINFORMATION);

        return true;
    }

    // 改进的递归删除注册表树方法
    bool DeleteRegistryTree(HKEY hParentKey, const wchar_t *subkey)
    {
        // 首先尝试使用 SHDeleteKeyW，它更可靠
        LONG result = SHDeleteKeyW(hParentKey, subkey);
        if (result == ERROR_SUCCESS || result == ERROR_FILE_NOT_FOUND)
        {
            return true;
        }

        // 如果 SHDeleteKeyW 失败，回退到手动删除
        HKEY hKey;
        result = RegOpenKeyExW(hParentKey, subkey, 0, KEY_READ | KEY_WRITE, &hKey);

        if (result != ERROR_SUCCESS)
        {
            // 如果键不存在，认为删除成功
            if (result == ERROR_FILE_NOT_FOUND)
            {
                return true;
            }
            return false;
        }

        // 枚举并删除所有子键
        DWORD index = 0;
        wchar_t childKeyName[256];
        DWORD childKeySize = sizeof(childKeyName) / sizeof(wchar_t);

        // 注意：删除子键时索引会变化，所以总是从0开始枚举
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

            // 重置大小
            childKeySize = sizeof(childKeyName) / sizeof(wchar_t);
        }

        RegCloseKey(hKey);

        // 删除当前键
        result = RegDeleteKeyW(hParentKey, subkey);
        return (result == ERROR_SUCCESS) || (result == ERROR_FILE_NOT_FOUND);
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
        // 取消编辑状态（如果有）
        if (isEditing)
        {
            CancelEditing();
        }

        // 显示正在刷新的提示
        MessageBoxW(hMainWindow, L"正在从注册表重新加载菜单项...", L"刷新", MB_OK | MB_ICONINFORMATION);

        // 强制从注册表重新读取所有菜单项
        ForceReloadFromRegistry();

        // 显示结果统计
        wchar_t resultMsg[256];
        swprintf(resultMsg, 256, L"刷新完成！\n总共找到 %d 个菜单项\n其中 %d 个是本程序创建的",
                 (int)allApps.size(),
                 (int)std::count_if(allApps.begin(), allApps.end(), [](const AppEntry &app)
                                    { return app.isCustom; }));

        MessageBoxW(hMainWindow, resultMsg, L"刷新完成", MB_OK | MB_ICONINFORMATION);
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
            else if (HIWORD(wParam) == LBN_DBLCLK && LOWORD(wParam) == 1001)
            { // 列表框双击事件
                OnListBoxDoubleClick();
            }
            else if (LOWORD(wParam) == 1101)
            { // 上下文菜单：在注册表中打开
                if (contextMenuIndex >= 0 && contextMenuIndex < (int)apps.size())
                {
                    OpenRegistryLocation(contextMenuIndex);
                }
            }
            else if (LOWORD(wParam) == 1102)
            { // 上下文菜单：刷新此项
                if (contextMenuIndex >= 0 && contextMenuIndex < (int)apps.size())
                {
                    RefreshSingleItemFromRegistry(apps[contextMenuIndex].name);
                    MessageBoxW(hMainWindow, L"已刷新选中项！", L"刷新", MB_OK | MB_ICONINFORMATION);
                }
            }
            break;

        case WM_SIZE:
            // 窗口大小改变时，重新计算水平滚动范围
            if (hListBox && !apps.empty())
            {
                // 延迟重新计算，确保布局已完成
                SetTimer(hMainWindow, 1001, 100, NULL); // 100ms 后重新计算
            }
            break;

        case WM_TIMER:
            if (wParam == 1001)
            {
                KillTimer(hMainWindow, 1001);
                // 重新计算水平滚动范围
                HDC hdc = GetDC(hListBox);
                if (hdc)
                {
                    int maxWidth = 0;
                    HFONT hOldFont = (HFONT)SelectObject(hdc, hModernFont);

                    for (const auto &app : apps)
                    {
                        std::wstring listText = GetDisplayText(app);
                        SIZE size;
                        if (GetTextExtentPoint32W(hdc, listText.c_str(), (int)listText.length(), &size))
                        {
                            if (size.cx > maxWidth)
                            {
                                maxWidth = size.cx;
                            }
                        }
                    }

                    SelectObject(hdc, hOldFont);
                    ReleaseDC(hListBox, hdc);

                    // 计算DPI缩放并设置水平滚动范围
                    HDC hdcWindow = GetDC(hMainWindow);
                    int dpiX = GetDeviceCaps(hdcWindow, LOGPIXELSX);
                    ReleaseDC(hMainWindow, hdcWindow);
                    float scale = dpiX / 96.0f;

                    int horizontalExtent = maxWidth + (int)(100 * scale);
                    SendMessage(hListBox, LB_SETHORIZONTALEXTENT, horizontalExtent, 0);
                }
            }
            break;

        case WM_CONTEXTMENU:
            // 处理右键菜单
            if ((HWND)wParam == hListBox)
            {
                int x = GET_X_LPARAM(lParam);
                int y = GET_Y_LPARAM(lParam);

                // 获取点击位置的项索引
                POINT pt = {x, y};
                ScreenToClient(hListBox, &pt);

                int index = (int)SendMessageW(hListBox, LB_ITEMFROMPOINT, 0, MAKELPARAM(pt.x, pt.y));
                if (HIWORD(index) == 0 && LOWORD(index) != LB_ERR) // 确保在项范围内
                {
                    int itemIndex = LOWORD(index);
                    if (itemIndex >= 0 && itemIndex < (int)apps.size())
                    {
                        // 选中该项
                        SendMessageW(hListBox, LB_SETCURSEL, itemIndex, 0);
                        ShowContextMenu(pt.x, pt.y, itemIndex);
                    }
                }
            }
            break;

        case WM_RBUTTONDOWN:
            // 处理右键点击（备用方法）
            if (hwnd == hListBox)
            {
                int x = GET_X_LPARAM(lParam);
                int y = GET_Y_LPARAM(lParam);

                int index = (int)SendMessageW(hListBox, LB_ITEMFROMPOINT, 0, MAKELPARAM(x, y));
                if (HIWORD(index) == 0 && LOWORD(index) != LB_ERR)
                {
                    int itemIndex = LOWORD(index);
                    if (itemIndex >= 0 && itemIndex < (int)apps.size())
                    {
                        // 选中该项
                        SendMessageW(hListBox, LB_SETCURSEL, itemIndex, 0);
                        ShowContextMenu(x, y, itemIndex);
                    }
                }
            }
            break;

        case WM_GETMINMAXINFO:
        {
            // 获取DPI缩放因子
            HDC hdc = GetDC(hMainWindow);
            int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
            ReleaseDC(hMainWindow, hdc);
            float scale = dpiX / 96.0f;

            // 根据DPI缩放锁定窗口大小
            MINMAXINFO *mmi = (MINMAXINFO *)lParam;
            mmi->ptMinTrackSize.x = (int)(750 * scale);
            mmi->ptMinTrackSize.y = (int)(500 * scale);
            mmi->ptMaxTrackSize.x = (int)(750 * scale);
            mmi->ptMaxTrackSize.y = (int)(500 * scale);
        }
        break;

        case WM_DESTROY:
            if (hContextMenu)
            {
                DestroyMenu(hContextMenu);
                hContextMenu = NULL;
            }
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

    // 设置DPI感知的兼容性方法
    HMODULE hUser32 = LoadLibraryW(L"user32.dll");
    HMODULE hShCore = LoadLibraryW(L"Shcore.dll");

    if (hShCore)
    {
        // 尝试使用 SetProcessDpiAwareness (需要Windows 8.1及以上)
        typedef HRESULT(WINAPI * FnSetProcessDpiAwareness)(PROCESS_DPI_AWARENESS);
        FnSetProcessDpiAwareness pSetProcessDpiAwareness =
            (FnSetProcessDpiAwareness)GetProcAddress(hShCore, "SetProcessDpiAwareness");

        if (pSetProcessDpiAwareness)
        {
            // 尝试设置为每监视器DPI感知
            pSetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);
        }
        FreeLibrary(hShCore);
    }
    else if (hUser32)
    {
        // 回退方案: 尝试使用旧版的 SetProcessDPIAware (Vista及以上系统支持)
        auto pSetProcessDPIAware = (BOOL(WINAPI *)())GetProcAddress(hUser32, "SetProcessDPIAware");
        if (pSetProcessDPIAware)
        {
            pSetProcessDPIAware();
        }
        FreeLibrary(hUser32);
    }

    // 检查管理员权限
    BOOL isAdmin = FALSE;
    HANDLE hToken = NULL;

    if (OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &hToken))
    {
        TOKEN_ELEVATION elevation;
        DWORD dwSize;

        if (GetTokenInformation(hToken, TokenElevation, &elevation, sizeof(elevation), &dwSize))
        {
            isAdmin = elevation.TokenIsElevated;
        }
        CloseHandle(hToken);
    }

    if (!isAdmin)
    {
        MessageBoxW(NULL,
                    L"此程序需要管理员权限才能修改注册表。\n请以管理员身份重新运行。",
                    L"权限不足",
                    MB_OK | MB_ICONWARNING);
        return 1;
    }

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