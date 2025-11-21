#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0A00 // Define as Windows 10
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

// Explicitly link required libraries
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "Shcore.lib")

// Application structure
struct AppEntry
{
    std::wstring name;        // Registry key name
    std::wstring path;        // Program path
    std::wstring displayName; // Display name
    std::wstring icon;        // Icon path
    bool isCustom;            // Whether created by this program
};

class RightClickManager
{
private:
    std::vector<AppEntry> apps;
    std::vector<AppEntry> allApps; // Store all apps for filtering
    HWND hMainWindow;
    HWND hListBox;
    HWND hAddButton;
    HWND hRemoveButton;
    HWND hRefreshButton;
    HWND hShowAllCheckbox;
    HWND hEditBox; // Edit box handle
    HANDLE hMutex;
    bool showAllItems;         // Whether to show all items
    bool isEditing;            // Whether in editing state
    HFONT hModernFont;         // Font handle
    int editingIndex;          // Index of item being edited
    WNDPROC oldEditProc;       // Original edit box procedure
    HMENU hContextMenu;        // Context menu handle
    int contextMenuIndex;      // Index for context menu
    bool isChineseOS;          // Whether the OS is Chinese
    bool forceEnglish = false; // 调试用：强制英文

    // Get localized string based on OS language
    std::wstring GetLocalizedString(const std::wstring &chinese, const std::wstring &english)
    {
        if (forceEnglish)
        {
            return english;
        }
        return isChineseOS ? chinese : english;
    }

    // 修改 IsChineseOS 函数，添加更严格的检测
    bool IsChineseOS()
    {
        LANGID langID = GetUserDefaultUILanguage();

        // 更严格的检测：只有简体中文和繁体中文才认为是中文系统
        return (PRIMARYLANGID(langID) == LANG_CHINESE &&
                (SUBLANGID(langID) == SUBLANG_CHINESE_SIMPLIFIED ||
                 SUBLANGID(langID) == SUBLANG_CHINESE_TRADITIONAL));
    }

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
                // Auto-save when losing focus
                pThis->FinishEditing(true);
                return 0;
            }
        }

        // Call original window procedure
        if (pThis && pThis->oldEditProc)
        {
            return CallWindowProc(pThis->oldEditProc, hwnd, uMsg, wParam, lParam);
        }

        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }

    // Sort apps alphabetically
    void SortAppsAlphabetically()
    {
        std::sort(allApps.begin(), allApps.end(), [](const AppEntry &a, const AppEntry &b)
                  {
            // Case-insensitive comparison
            return _wcsicmp(a.displayName.c_str(), b.displayName.c_str()) < 0; });
    }

    // Refresh single item display name from registry
    void RefreshSingleItemFromRegistry(const std::wstring &itemName)
    {
        // Find corresponding item in allApps
        for (auto &app : allApps)
        {
            if (app.name == itemName)
            {
                // Re-read display name from registry
                std::wstring displayPath = L"Directory\\Background\\shell\\";
                displayPath += itemName;

                HKEY hDisplayKey;
                if (RegOpenKeyExW(HKEY_CLASSES_ROOT, displayPath.c_str(), 0, KEY_READ, &hDisplayKey) == ERROR_SUCCESS)
                {
                    wchar_t displayName[256];
                    DWORD nameSize = sizeof(displayName);
                    if (RegQueryValueExW(hDisplayKey, NULL, NULL, NULL, (LPBYTE)displayName, &nameSize) == ERROR_SUCCESS)
                    {
                        // Update display name in memory
                        app.displayName = displayName;
                    }
                    RegCloseKey(hDisplayKey);
                }
                break;
            }
        }

        // Also update corresponding item in apps
        for (auto &app : apps)
        {
            if (app.name == itemName)
            {
                // Re-read display name from registry
                std::wstring displayPath = L"Directory\\Background\\shell\\";
                displayPath += itemName;

                HKEY hDisplayKey;
                if (RegOpenKeyExW(HKEY_CLASSES_ROOT, displayPath.c_str(), 0, KEY_READ, &hDisplayKey) == ERROR_SUCCESS)
                {
                    wchar_t displayName[256];
                    DWORD nameSize = sizeof(displayName);
                    if (RegQueryValueExW(hDisplayKey, NULL, NULL, NULL, (LPBYTE)displayName, &nameSize) == ERROR_SUCCESS)
                    {
                        // Update display name in memory
                        app.displayName = displayName;
                    }
                    RegCloseKey(hDisplayKey);
                }
                break;
            }
        }

        // Refresh list display
        FilterApps();
    }

    // System built-in context menu items blacklist
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
    // Generate sorted registry key name
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
                // 生成带语言标识和排序数字的注册表项名称 - 修复语言后缀
                wchar_t keyName[256];
                std::wstring langSuffix = isChineseOS ? L"_zh" : L"_en"; // 确保使用下划线前缀
                swprintf(keyName, 256, L"%02d_CustomApp_%s%s", position, displayName.c_str(), langSuffix.c_str());
                return keyName;
            }
            position++;
        }

        // 默认回退
        std::wstring langSuffix = isChineseOS ? L"_zh" : L"_en";
        return L"99_CustomApp_" + displayName + langSuffix;
    }

    // Reorder all registry items - improved version
    void ReorderRegistryItems()
    {
        // Only reorder custom apps
        std::vector<AppEntry> customApps;
        for (const auto &app : allApps)
        {
            if (app.isCustom)
            {
                customApps.push_back(app);
            }
        }

        // Sort by display name
        std::sort(customApps.begin(), customApps.end(), [](const AppEntry &a, const AppEntry &b)
                  { return _wcsicmp(a.displayName.c_str(), b.displayName.c_str()) < 0; });

        // Regenerate registry key names for all custom items
        int position = 1;
        std::vector<std::pair<std::wstring, std::wstring>> keyMappings; // <oldKey, newKey>

        for (const auto &app : customApps)
        {
            // Generate new registry key name with language suffix - 修复这里
            wchar_t newKeyName[256];
            std::wstring langSuffix = isChineseOS ? L"_zh" : L"_en";
            swprintf(newKeyName, 256, L"%02d_CustomApp_%s%s", position, app.displayName.c_str(), langSuffix.c_str());

            // Only process if name actually changed
            if (app.name != newKeyName)
            {
                keyMappings.push_back({app.name, newKeyName});
            }
            position++;
        }

        // Return directly if no items need renaming
        if (keyMappings.empty())
        {
            return;
        }

        // Create new keys and copy data for each item that needs renaming
        for (const auto &mapping : keyMappings)
        {
            const std::wstring &oldKeyName = mapping.first;
            const std::wstring &newKeyName = mapping.second;

            // Find corresponding app data
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

            // Create new registry key
            std::wstring newShellKey = L"Directory\\Background\\shell\\";
            newShellKey += newKeyName;

            HKEY hNewKey;
            if (RegCreateKeyExW(HKEY_CLASSES_ROOT, newShellKey.c_str(), 0, NULL, 0, KEY_WRITE, NULL, &hNewKey, NULL) == ERROR_SUCCESS)
            {
                // Copy display name - read directly from memory to ensure consistency
                RegSetValueExW(hNewKey, NULL, 0, REG_SZ,
                               (const BYTE *)app->displayName.c_str(),
                               (app->displayName.length() + 1) * sizeof(wchar_t));

                // Copy icon
                if (!app->icon.empty())
                {
                    RegSetValueExW(hNewKey, L"Icon", 0, REG_SZ,
                                   (const BYTE *)app->icon.c_str(),
                                   (app->icon.length() + 1) * sizeof(wchar_t));
                }

                RegCloseKey(hNewKey);

                // Create command subkey
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

                    // Update name in memory
                    app->name = newKeyName;
                }
                else
                {
                    // If command creation fails, delete shell key
                    RegDeleteKeyW(HKEY_CLASSES_ROOT, newShellKey.c_str());
                    continue;
                }
            }
            else
            {
                continue;
            }
            // 在 ReorderRegistryItems() 函数的最后，更新所有自定义应用的 isCustom 标志
            for (auto &app : allApps)
            {
                if (app.isCustom) // 如果原本就是自定义的
                {
                    // 重新检查并更新 isCustom 标志
                    std::wstring currentLangSuffix = isChineseOS ? L"_zh" : L"_en";
                    app.isCustom = (wcsstr(app.name.c_str(), L"CustomApp_") != nullptr) &&
                                   (wcsstr(app.name.c_str(), currentLangSuffix.c_str()) != nullptr);
                }
            }
        }

        // Delete all old registry keys (only delete old keys that successfully created new keys)
        for (const auto &mapping : keyMappings)
        {
            std::wstring oldShellKey = L"Directory\\Background\\shell\\";
            oldShellKey += mapping.first;

            // Ensure new key exists before deleting old key
            std::wstring newShellKey = L"Directory\\Background\\shell\\";
            newShellKey += mapping.second;

            HKEY hTestKey;
            if (RegOpenKeyExW(HKEY_CLASSES_ROOT, newShellKey.c_str(), 0, KEY_READ, &hTestKey) == ERROR_SUCCESS)
            {
                RegCloseKey(hTestKey);
                // New key exists, safe to delete old key
                DeleteRegistryTree(HKEY_CLASSES_ROOT, oldShellKey.c_str());
            }
        }

        // Refresh system
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

        // Reload all menu items - re-read from registry to ensure data consistency
        LoadAllContextMenuItems();
    }

    // Refresh isCustom flag for a single item
    void RefreshIsCustomFlag(AppEntry &app)
    {
        std::wstring currentLangSuffix = isChineseOS ? L"_zh" : L"_en";
        app.isCustom = (wcsstr(app.name.c_str(), L"CustomApp_") != nullptr) &&
                       (wcsstr(app.name.c_str(), currentLangSuffix.c_str()) != nullptr);
    }

    // Force reload all menu items from registry
    void ForceReloadFromRegistry()
    {
        // Clear existing data
        allApps.clear();
        apps.clear();

        // Reset list box
        SendMessageW(hListBox, LB_RESETCONTENT, 0, 0);

        // Re-read directly from registry
        HKEY hKey;
        if (RegOpenKeyExW(HKEY_CLASSES_ROOT, L"Directory\\Background\\shell", 0, KEY_READ, &hKey) == ERROR_SUCCESS)
        {
            wchar_t subkeyName[256];
            DWORD index = 0;
            DWORD nameSize = sizeof(subkeyName) / sizeof(wchar_t);

            while (RegEnumKeyExW(hKey, index, subkeyName, &nameSize, NULL, NULL, NULL, NULL) == ERROR_SUCCESS)
            {
                // Skip system items
                if (!IsSystemItem(subkeyName))
                {
                    AppEntry app;
                    app.name = subkeyName;
                    // Modify custom app identification logic to support new sorting naming format
                    app.isCustom = (wcsstr(subkeyName, L"CustomApp_") != nullptr);

                    // Get display name
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

                        // Get icon
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

                    // Get program path
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

        // Re-sort and filter app list
        SortAppsAlphabetically();
        FilterApps();
    }

public:
    RightClickManager() : hMainWindow(NULL), hListBox(NULL), hAddButton(NULL),
                          hRemoveButton(NULL), hRefreshButton(NULL), hShowAllCheckbox(NULL),
                          hEditBox(NULL), hMutex(NULL), showAllItems(false), isEditing(false),
                          hModernFont(NULL), editingIndex(-1), oldEditProc(NULL),
                          hContextMenu(NULL), contextMenuIndex(-1), isChineseOS(false) {}

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

    // 强制设置语言（用于测试）
    void ForceLanguage(bool chinese)
    {
        isChineseOS = chinese;
        // 重新创建控件以应用语言更改
        if (hMainWindow)
        {
            // 销毁现有控件
            if (hListBox)
                DestroyWindow(hListBox);
            if (hAddButton)
                DestroyWindow(hAddButton);
            if (hRemoveButton)
                DestroyWindow(hRemoveButton);
            if (hRefreshButton)
                DestroyWindow(hRefreshButton);
            if (hShowAllCheckbox)
                DestroyWindow(hShowAllCheckbox);

            // 重新创建控件
            CreateControls((HINSTANCE)GetWindowLongPtr(hMainWindow, GWLP_HINSTANCE));
            LoadAllContextMenuItems();
        }
    }

    // Create context menu
    void CreateContextMenu()
    {
        hContextMenu = CreatePopupMenu();
        if (hContextMenu)
        {
            AppendMenuW(hContextMenu, MF_STRING, 1101, GetLocalizedString(L"📁 在注册表中打开", L"📁 Open in Registry").c_str());
            AppendMenuW(hContextMenu, MF_SEPARATOR, 0, NULL);
            AppendMenuW(hContextMenu, MF_STRING, 1102, GetLocalizedString(L"🔄 刷新此项", L"🔄 Refresh Item").c_str());
        }
    }

    // Show context menu
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
            // Get list item rectangle position
            RECT itemRect;
            if (SendMessageW(hListBox, LB_GETITEMRECT, index, (LPARAM)&itemRect) != LB_ERR)
            {
                // Calculate menu display position (below the item)
                POINT pt;
                pt.x = itemRect.left + 10; // Slight offset to avoid covering text
                pt.y = itemRect.bottom;

                // Convert to screen coordinates
                ClientToScreen(hListBox, &pt);

                // Show context menu
                TrackPopupMenuEx(hContextMenu,
                                 TPM_RIGHTBUTTON | TPM_LEFTALIGN | TPM_TOPALIGN,
                                 pt.x, pt.y, hMainWindow, NULL);
            }
            else
            {
                // Alternative method: use mouse position
                POINT pt = {x, y};
                ClientToScreen(hListBox, &pt);
                TrackPopupMenuEx(hContextMenu,
                                 TPM_RIGHTBUTTON | TPM_LEFTALIGN | TPM_TOPALIGN,
                                 pt.x, pt.y, hMainWindow, NULL);
            }
        }
    }

    // Open registry location
    void OpenRegistryLocation(int index)
    {
        if (index < 0 || index >= (int)apps.size())
            return;

        AppEntry &app = apps[index];

        // Build registry path
        std::wstring regPath = L"计算机\\HKEY_CLASSES_ROOT\\Directory\\Background\\shell\\" + app.name;

        // Try to open Registry Editor using ShellExecute
        SHELLEXECUTEINFOW sei = {sizeof(sei)};
        sei.lpVerb = L"open";
        sei.lpFile = L"regedit.exe";
        sei.lpParameters = L""; // Don't use silent mode so user can see interface
        sei.nShow = SW_SHOW;

        if (ShellExecuteExW(&sei))
        {
            // Give Registry Editor some time to start
            Sleep(1000);

            // Try to send keys to Registry Editor to navigate to our path
            HWND hRegEdit = FindWindowW(L"RegEdit_RegEdit", NULL);
            if (hRegEdit)
            {
                // Activate window
                SetForegroundWindow(hRegEdit);
                Sleep(100);

                // Send Ctrl+F to open find dialog
                keybd_event(VK_CONTROL, 0, 0, 0);
                keybd_event('F', 0, 0, 0);
                keybd_event('F', 0, KEYEVENTF_KEYUP, 0);
                keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);

                // Wait for find dialog to open
                Sleep(500);

                // Enter our path in the find box
                // First ensure input method is in English state
                HWND hForeground = GetForegroundWindow();

                // Use more reliable way to input path
                for (wchar_t c : regPath)
                {
                    // For backslash, use virtual key code instead of character to avoid input method interference
                    if (c == L'\\')
                    {
                        // Directly send backslash virtual key code
                        keybd_event(VK_OEM_5, 0, 0, 0);
                        keybd_event(VK_OEM_5, 0, KEYEVENTF_KEYUP, 0);
                    }
                    else if (c == L'：') // Convert Chinese colon to English colon
                    {
                        keybd_event(VK_SHIFT, 0, 0, 0);
                        keybd_event(VK_OEM_1, 0, 0, 0);
                        keybd_event(VK_OEM_1, 0, KEYEVENTF_KEYUP, 0);
                        keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
                    }
                    else
                    {
                        // For other characters, use more reliable input method
                        SHORT vk = VkKeyScanW(c);
                        if (vk != -1)
                        {
                            BYTE vkCode = LOBYTE(vk);
                            BYTE shiftState = HIBYTE(vk);

                            if (shiftState & 1) // Shift key
                            {
                                keybd_event(VK_SHIFT, 0, 0, 0);
                            }
                            if (shiftState & 2) // Ctrl key
                            {
                                keybd_event(VK_CONTROL, 0, 0, 0);
                            }
                            if (shiftState & 4) // Alt key
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
                    Sleep(10); // Small delay to ensure characters are entered correctly
                }

                // Send Enter to start search
                Sleep(100);
                keybd_event(VK_RETURN, 0, 0, 0);
                keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0);

                std::wstring message = GetLocalizedString(
                    L"注册表编辑器已打开并尝试定位到指定位置。\n如果未自动定位，请手动导航到：\n" + regPath,
                    L"Registry Editor has been opened and attempted to locate the specified location.\nIf not automatically located, please navigate to:\n" + regPath);

                MessageBoxW(hMainWindow,
                            message.c_str(),
                            GetLocalizedString(L"打开注册表位置", L"Open Registry Location").c_str(),
                            MB_OK | MB_ICONINFORMATION);
            }
            else
            {
                std::wstring message = GetLocalizedString(
                    L"注册表编辑器已打开，但无法自动定位。\n请手动导航到：\n" + regPath,
                    L"Registry Editor has been opened but cannot automatically locate.\nPlease navigate manually to:\n" + regPath);

                MessageBoxW(hMainWindow,
                            message.c_str(),
                            GetLocalizedString(L"打开注册表位置", L"Open Registry Location").c_str(),
                            MB_OK | MB_ICONINFORMATION);
            }
        }
        else
        {
            // Alternative method: show path for manual navigation
            std::wstring message = GetLocalizedString(
                L"请手动在注册表编辑器中导航到以下路径：\n\n" + regPath,
                L"Please manually navigate to the following path in Registry Editor:\n\n" + regPath);

            MessageBoxW(hMainWindow,
                        message.c_str(),
                        GetLocalizedString(L"注册表位置", L"Registry Location").c_str(),
                        MB_OK | MB_ICONINFORMATION);
        }
    }

    // Handle list box double-click event
    void OnListBoxDoubleClick()
    {
        if (isEditing)
            return; // Ignore double-click if editing

        int selectedIndex = (int)SendMessageW(hListBox, LB_GETCURSEL, 0, 0);
        if (selectedIndex == LB_ERR)
            return;

        // Only allow renaming items created by this program
        if (selectedIndex >= 0 && selectedIndex < (int)apps.size() && apps[selectedIndex].isCustom)
        {
            StartEditing(selectedIndex);
        }
        else
        {
            MessageBoxW(hMainWindow,
                        GetLocalizedString(L"只能重命名本程序创建的项目（✅ 标记的项）", L"Only items created by this program (marked with ✅) can be renamed.").c_str(),
                        GetLocalizedString(L"提示", L"Information").c_str(),
                        MB_OK | MB_ICONINFORMATION);
        }
    }

    // Start editing
    void StartEditing(int index)
    {
        editingIndex = index;
        isEditing = true;

        // Get list item position
        RECT itemRect;
        SendMessageW(hListBox, LB_GETITEMRECT, index, (LPARAM)&itemRect);

        // Adjust edit box position to avoid icon area
        itemRect.left += 30; // Leave space for icon

        // Create edit box
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
            SendMessage(hEditBox, EM_SETSEL, 0, -1); // Select all text

            // Subclass edit box to capture keyboard messages
            SetWindowLongPtr(hEditBox, GWLP_USERDATA, (LONG_PTR)this);
            oldEditProc = (WNDPROC)SetWindowLongPtr(hEditBox, GWLP_WNDPROC, (LONG_PTR)EditBoxProc);
        }
    }

    // Finish editing - improved version
    void FinishEditing(bool saveChanges)
    {
        if (!isEditing)
            return;

        if (saveChanges && hEditBox)
        {
            // Get new name from edit box
            wchar_t newName[256];
            GetWindowTextW(hEditBox, newName, 256);

            if (wcslen(newName) > 0)
            {
                AppEntry &app = apps[editingIndex];
                std::wstring oldDisplayName = app.displayName;

                // Check if name actually changed
                if (lstrcmpiW(oldDisplayName.c_str(), newName) != 0)
                {
                    // Update display name in registry
                    std::wstring shellKey = L"Directory\\Background\\shell\\";
                    shellKey += app.name;

                    HKEY hKey;
                    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, shellKey.c_str(), 0, KEY_WRITE, &hKey) == ERROR_SUCCESS)
                    {
                        RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE *)newName, (wcslen(newName) + 1) * sizeof(wchar_t));
                        RegCloseKey(hKey);

                        // Refresh system
                        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

                        // Re-read display name directly from registry for that item to ensure data synchronization
                        RefreshSingleItemFromRegistry(app.name);

                        // Reorder registry items
                        ReorderRegistryItems();

                        MessageBoxW(hMainWindow,
                                    GetLocalizedString(L"重命名成功！如果菜单项消失请刷新", L"Rename successful! If a menu item disappears, please refresh").c_str(),
                                    GetLocalizedString(L"成功", L"Success").c_str(),
                                    MB_OK | MB_ICONINFORMATION);
                    }
                    else
                    {
                        MessageBoxW(hMainWindow,
                                    GetLocalizedString(L"重命名失败！请以管理员身份运行。", L"Rename failed! Please run as administrator.").c_str(),
                                    GetLocalizedString(L"错误", L"Error").c_str(),
                                    MB_OK | MB_ICONERROR);
                    }
                }
                else
                {
                    // Name didn't change, no operation needed
                    MessageBoxW(hMainWindow,
                                GetLocalizedString(L"名称没有改变。", L"The name has not changed.").c_str(),
                                GetLocalizedString(L"提示", L"Information").c_str(),
                                MB_OK | MB_ICONINFORMATION);
                }
            }
        }

        // Clean up editing state
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

    // Cancel editing
    void CancelEditing()
    {
        FinishEditing(false);
    }

    // Check if another instance is already running
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

    // Activate existing instance window
    void ActivateExistingInstance()
    {
        HWND hExistingWindow = FindWindowW(L"RightClickManager", GetLocalizedString(L"桌面右键菜单管理器", L"Desktop Context Menu Manager").c_str());
        if (hExistingWindow)
        {
            if (IsIconic(hExistingWindow))
            {
                ShowWindow(hExistingWindow, SW_RESTORE);
            }
            SetForegroundWindow(hExistingWindow);
        }
    }

    // Initialize program
    bool Initialize(HINSTANCE hInstance)
    {
        // 检测OS语言
        isChineseOS = IsChineseOS();

        // 调试：通过命令行参数强制英文
        LPWSTR cmdLine = GetCommandLineW();
        if (wcsstr(cmdLine, L"-english") != NULL ||
            wcsstr(cmdLine, L"/english") != NULL ||
            wcsstr(cmdLine, L"-en") != NULL)
        {
            forceEnglish = true;
            isChineseOS = false;
        }

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

        // 根据语言计算窗口大小
        int windowWidth, windowHeight;
        if (isChineseOS)
        {
            windowWidth = (int)(750 * scale);
            windowHeight = (int)(500 * scale);
        }
        else
        {
            windowWidth = (int)(850 * scale); // 英文界面更宽
            windowHeight = (int)(500 * scale);
        }

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

        // 计算窗口位置使其居中
        int screenWidth = GetSystemMetrics(SM_CXSCREEN);
        int screenHeight = GetSystemMetrics(SM_CYSCREEN);
        int x = (screenWidth - windowWidth) / 2;
        int y = (screenHeight - windowHeight) / 2;

        // 创建窗口时直接使用正确的大小
        hMainWindow = CreateWindowExW(
            WS_EX_APPWINDOW,
            L"RightClickManager",
            GetLocalizedString(L"桌面右键菜单管理器", L"Desktop Context Menu Manager").c_str(),
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
            x, y,
            windowWidth, windowHeight, // 这里使用计算好的大小
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

    // Create controls
    void CreateControls(HINSTANCE hInstance)
    {
        // 获取DPI缩放因子
        HDC hdc = GetDC(hMainWindow);
        int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
        ReleaseDC(hMainWindow, hdc);
        float scale = dpiX / 96.0f;

        // 根据语言选择不同的布局参数
        int listBoxWidth, listBoxHeight, buttonWidth, buttonHeight;
        int margin, rightPanelX, helpTextWidth, helpTextHeight;

        if (isChineseOS)
        {
            // 中文界面：保持原有布局
            listBoxWidth = (int)(580 * scale);
            listBoxHeight = (int)(380 * scale);
            buttonWidth = (int)(140 * scale);
            buttonHeight = (int)(30 * scale);
            margin = (int)(10 * scale);
            rightPanelX = (int)(600 * scale);
            helpTextWidth = (int)(140 * scale);
            helpTextHeight = (int)(200 * scale);
        }
        else
        {
            // 英文界面：使用新布局
            listBoxWidth = (int)(580 * scale);
            listBoxHeight = (int)(380 * scale);
            buttonWidth = (int)(160 * scale); // 增加按钮宽度
            buttonHeight = (int)(30 * scale);
            margin = (int)(10 * scale);
            rightPanelX = (int)(620 * scale);    // 调整右侧面板位置
            helpTextWidth = (int)(200 * scale);  // 增加帮助文本宽度
            helpTextHeight = (int)(220 * scale); // 增加帮助文本高度
        }

        // 注意：这里不再调用 SetWindowPos，因为窗口大小已经在 Initialize 中设置正确

        // 列表控件
        hListBox = CreateWindowExW(
            WS_EX_CLIENTEDGE,
            L"LISTBOX",
            L"",
            WS_CHILD | WS_VISIBLE | LBS_NOTIFY | LBS_HASSTRINGS |
                WS_VSCROLL | WS_HSCROLL | LBS_NOINTEGRALHEIGHT | LBS_DISABLENOSCROLL,
            margin, margin,
            listBoxWidth, listBoxHeight,
            hMainWindow,
            (HMENU)1001,
            hInstance,
            NULL);

        // 设置列表项高度
        int itemHeight = (int)(24 * scale);
        SendMessage(hListBox, LB_SETITEMHEIGHT, 0, itemHeight);

        // 添加按钮
        hAddButton = CreateWindowW(
            L"BUTTON",
            GetLocalizedString(L"📁 添加程序", L"📁 Add Program").c_str(),
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
            GetLocalizedString(L"🗑️ 删除选中", L"🗑️ Remove Selected").c_str(),
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
            GetLocalizedString(L"🔄 刷新列表", L"🔄 Refresh List").c_str(),
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
            GetLocalizedString(L"显示所有项目", L"Show All Items").c_str(),
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
            rightPanelX, margin + (buttonHeight + margin / 2) * 3,
            buttonWidth, buttonHeight,
            hMainWindow,
            (HMENU)1005,
            hInstance,
            NULL);

        // 帮助文本 - 根据语言使用不同的文本
        HWND hHelpText = CreateWindowW(
            L"STATIC",
            GetLocalizedString(
                // 中文保持原样
                L"💡 桌面右键菜单管理\n\n"
                L"✅ 本程序创建的项目\n"
                L"📌 其他程序创建的项目\n\n"
                L"🖱️ 操作提示:\n"
                L"• 双击 ✅ 项可以重命名\n"
                L"• 右键项打开功能菜单\n"
                L"• 勾选复选框显示所有项目",

                // 英文使用简化文本
                L"💡 Context Menu Manager\n\n"
                L"✅ Items created by app\n"
                L"📌 Other program's items\n\n"
                L"🖱️ Tips:\n"
                L"• Double-click ✅ to rename\n"
                L"• Right-click for menu\n"
                L"• Check to show all")
                .c_str(),
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

        // Truncate if text is too long (this is just for display, full content can still be viewed via scrolling)
        if (baseText.length() > maxDisplayLength)
        {
            // Keep important information at beginning and end
            std::wstring shortened = baseText.substr(0, maxDisplayLength - 10) + L"..." +
                                     baseText.substr(baseText.length() - 7);
            return shortened;
        }

        return baseText;
    }

    // Check if system item
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

    // Load all context menu items
    void LoadAllContextMenuItems()
    {

        // If editing, cancel editing first
        if (isEditing)
        {
            CancelEditing();
        }

        // Clear existing data
        allApps.clear();
        apps.clear();
        SendMessageW(hListBox, LB_RESETCONTENT, 0, 0);

        // Check desktop context menu registry location
        HKEY hKey;
        if (RegOpenKeyExW(HKEY_CLASSES_ROOT, L"Directory\\Background\\shell", 0, KEY_READ, &hKey) == ERROR_SUCCESS)
        {
            wchar_t subkeyName[256];
            DWORD index = 0;
            DWORD nameSize = sizeof(subkeyName) / sizeof(wchar_t);

            while (RegEnumKeyExW(hKey, index, subkeyName, &nameSize, NULL, NULL, NULL, NULL) == ERROR_SUCCESS)
            {
                // Skip system items
                if (!IsSystemItem(subkeyName))
                {
                    AppEntry app;
                    app.name = subkeyName;
                    // 在加载每个应用项时，确保 isCustom 标志正确设置
                    std::wstring currentLangSuffix = isChineseOS ? L"_zh" : L"_en";
                    app.isCustom = (wcsstr(subkeyName, L"CustomApp_") != nullptr) &&
                                   (wcsstr(subkeyName, currentLangSuffix.c_str()) != nullptr);

                    // Get display name
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
                            app.displayName = subkeyName; // If no display name, use registry key name
                        }

                        // Get icon
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

                    // Get program path
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

        // Sort by display name alphabetically
        SortAppsAlphabetically();

        // Filter app list based on display settings
        FilterApps();
    }

    // Filter app list based on display settings
    void FilterApps()
    {
        // Cancel editing state if editing
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

                // Use optimized display text
                std::wstring listText = GetDisplayText(app);
                SendMessageW(hListBox, LB_ADDSTRING, 0, (LPARAM)listText.c_str());
            }
        }

        // Set horizontal scroll range so long text can be scrolled to view
        if (!apps.empty())
        {
            HDC hdc = GetDC(hListBox);
            if (hdc)
            {
                int maxWidth = 0;
                HFONT hOldFont = (HFONT)SelectObject(hdc, hModernFont); // Ensure correct font is used

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

                // Set horizontal scroll range, add sufficient margin
                // Calculate DPI scaling
                HDC hdcWindow = GetDC(hMainWindow);
                int dpiX = GetDeviceCaps(hdcWindow, LOGPIXELSX);
                ReleaseDC(hMainWindow, hdcWindow);
                float scale = dpiX / 96.0f;

                // Add more margin to ensure long text is fully visible
                int horizontalExtent = maxWidth + (int)(100 * scale);
                SendMessage(hListBox, LB_SETHORIZONTALEXTENT, horizontalExtent, 0);
            }
        }

        // Ensure vertical scrollbar displays correctly when too many items
        // Get list box client area height
        RECT listRect;
        GetClientRect(hListBox, &listRect);
        int listHeight = listRect.bottom - listRect.top;

        // Get list item height
        int itemHeight = (int)SendMessage(hListBox, LB_GETITEMHEIGHT, 0, 0);
        if (itemHeight <= 0)
            itemHeight = 20; // Default value

        // Calculate number of items that can fit in visible area
        int visibleItems = listHeight / itemHeight;

        // If number of items exceeds visible area, vertical scrollbar will automatically display
        // We can select the first item to give user visual feedback
        if (!apps.empty())
        {
            SendMessageW(hListBox, LB_SETCURSEL, 0, 0);
        }
    }

    // Clean app path
    void CleanAppPath(std::wstring &path)
    {
        // Remove quotes
        if (path.length() >= 2 && path[0] == L'\"' && path[path.length() - 1] == L'\"')
        {
            path = path.substr(1, path.length() - 2);
        }
        // Remove parameters
        size_t pos = path.find(L".exe");
        if (pos != std::wstring::npos)
        {
            path = path.substr(0, pos + 4);
        }
    }

    // Add app to desktop context menu
    bool AddAppToContextMenu(const std::wstring &appPath)
    {
        // Get program name
        size_t lastSlash = appPath.find_last_of(L'\\');
        size_t lastDot = appPath.find_last_of(L'.');
        if (lastSlash == std::wstring::npos)
            return false;

        std::wstring appName = appPath.substr(lastSlash + 1);
        if (lastDot != std::wstring::npos && lastDot > lastSlash)
        {
            appName = appPath.substr(lastSlash + 1, lastDot - lastSlash - 1);
        }

        // Generate registry key name with sort number
        std::wstring registryKey = GenerateSortedRegistryKey(appName);

        // Create registry key
        std::wstring shellKey = L"Directory\\Background\\shell\\";
        shellKey += registryKey;

        HKEY hKey;
        LONG result = RegCreateKeyExW(HKEY_CLASSES_ROOT, shellKey.c_str(), 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL);
        if (result == ERROR_SUCCESS)
        {
            // Set display name
            result = RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE *)appName.c_str(), (appName.length() + 1) * sizeof(wchar_t));
            if (result != ERROR_SUCCESS)
            {
                // Add error information
                wchar_t errorMsg[256];
                swprintf(errorMsg, 256, GetLocalizedString(L"设置显示名称失败！错误代码: %d", L"Failed to set display name! Error code: %d").c_str(), result);
                MessageBoxW(hMainWindow, errorMsg, GetLocalizedString(L"错误", L"Error").c_str(), MB_OK | MB_ICONERROR);
                RegCloseKey(hKey);
                return false;
            }

            // Set icon
            std::wstring iconValue = L"\"";
            iconValue += appPath;
            iconValue += L"\"";

            RegSetValueExW(hKey, L"Icon", 0, REG_SZ, (const BYTE *)iconValue.c_str(), (iconValue.length() + 1) * sizeof(wchar_t));

            RegCloseKey(hKey);
        }
        else
        {
            // Add error information
            wchar_t errorMsg[256];
            swprintf(errorMsg, 256, GetLocalizedString(L"创建注册表项失败！错误代码: %d", L"Failed to create registry key! Error code: %d").c_str(), result);
            MessageBoxW(hMainWindow, errorMsg, GetLocalizedString(L"错误", L"Error").c_str(), MB_OK | MB_ICONERROR);
            return false;
        }

        // Create command subkey
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
                // Refresh system to make registry changes take effect immediately
                SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

                // Reorder registry items - this will call LoadAllContextMenuItems()
                ReorderRegistryItems();
                return true;
            }
            else
            {
                // Add error information
                wchar_t errorMsg[256];
                swprintf(errorMsg, 256, GetLocalizedString(L"设置命令失败！错误代码: %d", L"Failed to set command! Error code: %d").c_str(), result);
                MessageBoxW(hMainWindow, errorMsg, GetLocalizedString(L"错误", L"Error").c_str(), MB_OK | MB_ICONERROR);
            }
        }
        else
        {
            // Add error information
            wchar_t errorMsg[256];
            swprintf(errorMsg, 256, GetLocalizedString(L"创建命令子键失败！错误代码: %d", L"Failed to create command subkey! Error code: %d").c_str(), GetLastError());
            MessageBoxW(hMainWindow, errorMsg, GetLocalizedString(L"错误", L"Error").c_str(), MB_OK | MB_ICONERROR);
        }

        return false;
    }

    // Remove app from context menu
    bool RemoveAppFromContextMenu(int index)
    {
        if (index < 0 || index >= (int)apps.size())
            return false;

        AppEntry &app = apps[index];

        // Save item information to be deleted
        std::wstring deleteName = app.name;
        std::wstring deleteDisplayName = app.displayName;
        std::wstring deletePath = app.path;

        // Show extra warning for items not created by this program
        if (!app.isCustom)
        {
            std::wstring warningMsg = GetLocalizedString(
                L"警告：此项不是由本程序创建，可能是系统或其他应用程序的右键菜单项。\n\n"
                L"名称: " +
                    app.displayName + L"\n"
                                      L"路径: " +
                    app.path + L"\n\n"
                               L"确定要删除吗？",

                L"Warning: This item was not created by this program and may be a system or other application's context menu item.\n\n"
                L"Name: " +
                    app.displayName + L"\n"
                                      L"Path: " +
                    app.path + L"\n\n"
                               L"Are you sure you want to delete it?");

            if (MessageBoxW(hMainWindow, warningMsg.c_str(), GetLocalizedString(L"确认删除系统项", L"Confirm Deletion of System Item").c_str(), MB_YESNO | MB_ICONWARNING) != IDYES)
            {
                return false;
            }
        }

        std::wstring shellKey = L"Directory\\Background\\shell\\";
        shellKey += app.name;

        // First try normal deletion
        bool deleteSuccess = DeleteRegistryTree(HKEY_CLASSES_ROOT, shellKey.c_str());

        if (!deleteSuccess)
        {
            // If normal deletion fails, show detailed error information
            DWORD errorCode = GetLastError();
            wchar_t errorMsg[512];
            swprintf(errorMsg, 512,
                     GetLocalizedString(
                         L"删除失败！错误代码: %d\n\n"
                         L"可能的原因：\n"
                         L"• 注册表项被其他进程占用\n"
                         L"• 权限不足\n"
                         L"• 注册表项不存在\n\n"
                         L"请尝试以管理员身份运行程序，或重启后重试。",

                         L"Deletion failed! Error code: %d\n\n"
                         L"Possible reasons:\n"
                         L"• Registry key is being used by another process\n"
                         L"• Insufficient permissions\n"
                         L"• Registry key does not exist\n\n"
                         L"Please try running as administrator, or restart and try again.")
                         .c_str(),
                     errorCode);

            MessageBoxW(hMainWindow, errorMsg, GetLocalizedString(L"删除失败", L"Deletion Failed").c_str(), MB_OK | MB_ICONERROR);
            return false;
        }

        // Refresh system
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

        // Reload directly from registry
        ForceReloadFromRegistry();

        // Show success message
        MessageBoxW(hMainWindow,
                    GetLocalizedString(
                        L"程序已从桌面右键菜单中删除！\n"
                        L"如果菜单项仍然显示，请尝试刷新桌面(F5)或重启资源管理器。",

                        L"Program has been removed from desktop context menu!\n"
                        L"If the menu item still appears, try refreshing desktop (F5) or restarting Explorer.")
                        .c_str(),
                    GetLocalizedString(L"删除成功", L"Deletion Successful").c_str(), MB_OK | MB_ICONINFORMATION);

        return true;
    }

    // Improved recursive registry tree deletion method
    bool DeleteRegistryTree(HKEY hParentKey, const wchar_t *subkey)
    {
        // First try using SHDeleteKeyW, it's more reliable
        LONG result = SHDeleteKeyW(hParentKey, subkey);
        if (result == ERROR_SUCCESS || result == ERROR_FILE_NOT_FOUND)
        {
            return true;
        }

        // If SHDeleteKeyW fails, fall back to manual deletion
        HKEY hKey;
        result = RegOpenKeyExW(hParentKey, subkey, 0, KEY_READ | KEY_WRITE, &hKey);

        if (result != ERROR_SUCCESS)
        {
            // If key doesn't exist, consider deletion successful
            if (result == ERROR_FILE_NOT_FOUND)
            {
                return true;
            }
            return false;
        }

        // Enumerate and delete all subkeys
        DWORD index = 0;
        wchar_t childKeyName[256];
        DWORD childKeySize = sizeof(childKeyName) / sizeof(wchar_t);

        // Note: When deleting subkeys, index changes, so always enumerate from 0
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

            // Reset size
            childKeySize = sizeof(childKeyName) / sizeof(wchar_t);
        }

        RegCloseKey(hKey);

        // Delete current key
        result = RegDeleteKeyW(hParentKey, subkey);
        return (result == ERROR_SUCCESS) || (result == ERROR_FILE_NOT_FOUND);
    }

    // Handle button clicks
    void OnAddButtonClick()
    {
        OPENFILENAMEW ofn;
        wchar_t fileName[MAX_PATH] = L"";

        ZeroMemory(&ofn, sizeof(ofn));
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner = hMainWindow;
        ofn.lpstrFile = fileName;
        ofn.nMaxFile = MAX_PATH;
        ofn.lpstrFilter = L"Executable files\0*.exe\0All files\0*.*\0";
        ofn.nFilterIndex = 1;
        ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

        if (GetOpenFileNameW(&ofn))
        {
            if (AddAppToContextMenu(fileName))
            {
                MessageBoxW(hMainWindow,
                            GetLocalizedString(
                                L"程序已成功添加到桌面右键菜单！\n"
                                L"程序图标也会显示在菜单中。\n"
                                L"可能需要刷新桌面或重新启动资源管理器才能看到变化。",

                                L"Program successfully added to desktop context menu!\n"
                                L"The program icon will also appear in the menu.\n"
                                L"You may need to refresh the desktop or restart Explorer to see the change.")
                                .c_str(),
                            GetLocalizedString(L"成功", L"Success").c_str(), MB_OK | MB_ICONINFORMATION);
            }
            else
            {
                MessageBoxW(hMainWindow,
                            GetLocalizedString(L"添加程序失败！请以管理员身份运行程序。", L"Failed to add program! Please run as administrator.").c_str(),
                            GetLocalizedString(L"错误", L"Error").c_str(), MB_OK | MB_ICONERROR);
            }
        }
    }

    void OnRemoveButtonClick()
    {
        int selectedIndex = (int)SendMessageW(hListBox, LB_GETCURSEL, 0, 0);
        if (selectedIndex == LB_ERR)
        {
            MessageBoxW(hMainWindow,
                        GetLocalizedString(L"请先选择一个程序！", L"Please select a program first!").c_str(),
                        GetLocalizedString(L"提示", L"Information").c_str(), MB_OK | MB_ICONINFORMATION);
            return;
        }

        AppEntry &app = apps[selectedIndex];
        std::wstring confirmMsg = GetLocalizedString(
            L"确定要从桌面右键菜单中删除这个程序吗？\n\n"
            L"名称: " +
                app.displayName + L"\n"
                                  L"路径: " +
                app.path,

            L"Are you sure you want to remove this program from the desktop context menu?\n\n"
            L"Name: " +
                app.displayName + L"\n"
                                  L"Path: " +
                app.path);

        if (MessageBoxW(hMainWindow, confirmMsg.c_str(), GetLocalizedString(L"确认删除", L"Confirm Deletion").c_str(), MB_YESNO | MB_ICONQUESTION) == IDYES)
        {
            if (RemoveAppFromContextMenu(selectedIndex))
            {
                MessageBoxW(hMainWindow,
                            GetLocalizedString(L"程序已从桌面右键菜单中删除！", L"Program has been removed from desktop context menu!").c_str(),
                            GetLocalizedString(L"成功", L"Success").c_str(), MB_OK | MB_ICONINFORMATION);
            }
            else
            {
                MessageBoxW(hMainWindow,
                            GetLocalizedString(L"删除程序失败！", L"Failed to remove program!").c_str(),
                            GetLocalizedString(L"错误", L"Error").c_str(), MB_OK | MB_ICONERROR);
            }
        }
    }

    void OnRefreshButtonClick()
    {
        // Cancel editing state (if any)
        if (isEditing)
        {
            CancelEditing();
        }

        // Show refreshing prompt
        MessageBoxW(hMainWindow,
                    GetLocalizedString(L"正在从注册表重新加载菜单项...", L"Reloading menu items from registry...").c_str(),
                    GetLocalizedString(L"刷新", L"Refresh").c_str(), MB_OK | MB_ICONINFORMATION);

        // Force reload all menu items from registry
        ForceReloadFromRegistry();

        // Show result statistics
        wchar_t resultMsg[256];
        swprintf(resultMsg, 256, GetLocalizedString(L"刷新完成！\n总共找到 %d 个菜单项\n其中 %d 个是本程序创建的", L"Refresh completed!\nTotal found: %d menu items\n%d items created by this program").c_str(),
                 (int)allApps.size(),
                 (int)std::count_if(allApps.begin(), allApps.end(), [](const AppEntry &app)
                                    { return app.isCustom; }));

        MessageBoxW(hMainWindow, resultMsg, GetLocalizedString(L"刷新完成", L"Refresh Completed").c_str(), MB_OK | MB_ICONINFORMATION);
    }

    void OnShowAllCheckboxClick()
    {
        showAllItems = (SendMessage(hShowAllCheckbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
        FilterApps();
    }

    // Static window procedure function
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

    // Message handling function
    LRESULT HandleMessage(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
    {
        switch (uMsg)
        {
        case WM_COMMAND:
            if (LOWORD(wParam) == 1002)
            { // Add button
                OnAddButtonClick();
            }
            else if (LOWORD(wParam) == 1003)
            { // Remove button
                OnRemoveButtonClick();
            }
            else if (LOWORD(wParam) == 1004)
            { // Refresh button
                OnRefreshButtonClick();
            }
            else if (LOWORD(wParam) == 1005)
            { // Checkbox
                OnShowAllCheckboxClick();
            }
            else if (HIWORD(wParam) == LBN_DBLCLK && LOWORD(wParam) == 1001)
            { // List box double-click event
                OnListBoxDoubleClick();
            }
            else if (LOWORD(wParam) == 1101)
            { // Context menu: Open in Registry
                if (contextMenuIndex >= 0 && contextMenuIndex < (int)apps.size())
                {
                    OpenRegistryLocation(contextMenuIndex);
                }
            }
            else if (LOWORD(wParam) == 1102)
            { // Context menu: Refresh item
                if (contextMenuIndex >= 0 && contextMenuIndex < (int)apps.size())
                {
                    RefreshSingleItemFromRegistry(apps[contextMenuIndex].name);
                    MessageBoxW(hMainWindow,
                                GetLocalizedString(L"已刷新选中项！", L"Selected item refreshed!").c_str(),
                                GetLocalizedString(L"刷新", L"Refresh").c_str(), MB_OK | MB_ICONINFORMATION);
                }
            }
            break;

        case WM_SIZE:
            // When window size changes, recalculate horizontal scroll range
            if (hListBox && !apps.empty())
            {
                // Delay recalculation to ensure layout is complete
                SetTimer(hMainWindow, 1001, 100, NULL); // Recalculate after 100ms
            }
            break;

        case WM_TIMER:
            if (wParam == 1001)
            {
                KillTimer(hMainWindow, 1001);
                // Recalculate horizontal scroll range
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

                    // Calculate DPI scaling and set horizontal scroll range
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
            // Handle right-click menu
            if ((HWND)wParam == hListBox)
            {
                int x = GET_X_LPARAM(lParam);
                int y = GET_Y_LPARAM(lParam);

                // Get item index at click position
                POINT pt = {x, y};
                ScreenToClient(hListBox, &pt);

                int index = (int)SendMessageW(hListBox, LB_ITEMFROMPOINT, 0, MAKELPARAM(pt.x, pt.y));
                if (HIWORD(index) == 0 && LOWORD(index) != LB_ERR) // Ensure within item range
                {
                    int itemIndex = LOWORD(index);
                    if (itemIndex >= 0 && itemIndex < (int)apps.size())
                    {
                        // Select the item
                        SendMessageW(hListBox, LB_SETCURSEL, itemIndex, 0);
                        ShowContextMenu(pt.x, pt.y, itemIndex);
                    }
                }
            }
            break;

        case WM_RBUTTONDOWN:
            // Handle right-click (alternative method)
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
                        // Select the item
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

            // 根据语言锁定窗口大小
            MINMAXINFO *mmi = (MINMAXINFO *)lParam;
            if (isChineseOS)
            {
                // 中文界面：保持原有大小限制
                mmi->ptMinTrackSize.x = (int)(750 * scale);
                mmi->ptMinTrackSize.y = (int)(500 * scale);
                mmi->ptMaxTrackSize.x = (int)(750 * scale);
                mmi->ptMaxTrackSize.y = (int)(500 * scale);
            }
            else
            {
                // 英文界面：增加宽度限制
                mmi->ptMinTrackSize.x = (int)(850 * scale);
                mmi->ptMinTrackSize.y = (int)(500 * scale);
                mmi->ptMaxTrackSize.x = (int)(850 * scale);
                mmi->ptMaxTrackSize.y = (int)(500 * scale);
            }
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

    // Run message loop
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

// Program entry point
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{

    // 全局语言强制 - 在创建 manager 实例之前
    bool globalForceEnglish = false;
    LPWSTR cmdLine = GetCommandLineW();

    if (wcsstr(cmdLine, L"-english") != NULL ||
        wcsstr(cmdLine, L"/english") != NULL ||
        wcsstr(cmdLine, L"-en") != NULL)
    {
        globalForceEnglish = true;

        // 设置线程UI语言
        SetThreadUILanguage(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US));
    }

    // Set DPI awareness compatibility method
    HMODULE hUser32 = LoadLibraryW(L"user32.dll");
    HMODULE hShCore = LoadLibraryW(L"Shcore.dll");

    if (hShCore)
    {
        // Try to use SetProcessDpiAwareness (requires Windows 8.1 or above)
        typedef HRESULT(WINAPI * FnSetProcessDpiAwareness)(PROCESS_DPI_AWARENESS);
        FnSetProcessDpiAwareness pSetProcessDpiAwareness =
            (FnSetProcessDpiAwareness)GetProcAddress(hShCore, "SetProcessDpiAwareness");

        if (pSetProcessDpiAwareness)
        {
            // Try to set as per-monitor DPI aware
            pSetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE);
        }
        FreeLibrary(hShCore);
    }
    else if (hUser32)
    {
        // Fallback: try to use older SetProcessDPIAware (supported on Vista and above)
        auto pSetProcessDPIAware = (BOOL(WINAPI *)())GetProcAddress(hUser32, "SetProcessDPIAware");
        if (pSetProcessDPIAware)
        {
            pSetProcessDPIAware();
        }
        FreeLibrary(hUser32);
    }

    // Check administrator privileges
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
        // Detect OS language for admin warning message
        bool isChinese = (PRIMARYLANGID(GetUserDefaultUILanguage()) == LANG_CHINESE);

        MessageBoxW(NULL,
                    isChinese ? L"此程序需要管理员权限才能修改注册表。\n请以管理员身份重新运行。" : L"This program requires administrator privileges to modify the registry.\nPlease run as administrator.",
                    isChinese ? L"权限不足" : L"Insufficient Privileges",
                    MB_OK | MB_ICONWARNING);
        return 1;
    }

    RightClickManager manager;

    // 如果全局强制英文，在这里设置
    if (globalForceEnglish)
    {
        // 我们需要在初始化后设置语言
        // 由于初始化在构造函数后调用，我们稍后在 Initialize 中处理
    }

    // Check if another instance is already running
    if (manager.IsAlreadyRunning())
    {
        manager.ActivateExistingInstance();
        return 0; // Exit directly, don't create new instance
    }

    if (!manager.Initialize(hInstance))
    {
        bool isChinese = (PRIMARYLANGID(GetUserDefaultUILanguage()) == LANG_CHINESE);
        MessageBoxW(NULL,
                    isChinese ? L"程序初始化失败！" : L"Program initialization failed!",
                    isChinese ? L"错误" : L"Error",
                    MB_OK | MB_ICONERROR);
        return 1;
    }

    return manager.Run();
}