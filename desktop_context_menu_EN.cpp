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
    HWND hMoveUpButton;   // Move up button
    HWND hMoveDownButton; // Move down button
    HWND hEditBox;        // Edit box handle
    HANDLE hMutex;
    bool showAllItems;    // Whether to show all items
    bool isEditing;       // Whether in editing state
    HFONT hModernFont;    // Font handle
    int editingIndex;     // Index of item being edited
    WNDPROC oldEditProc;  // Original edit box procedure
    HMENU hContextMenu;   // Context menu handle
    int contextMenuIndex; // Index of context menu item

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

    // Sort by registry key name (maintain actual order in registry)
    void SortAppsByRegistryKeyName()
    {
        std::sort(allApps.begin(), allApps.end(), [](const AppEntry &a, const AppEntry &b)
                  {
                  // Sort by registry key name, maintain original registry order
                  return _wcsicmp(a.name.c_str(), b.name.c_str()) < 0; });
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

    // System built-in right-click menu items blacklist
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
    // Generate registry key name - simplified version, no auto-sorting
    std::wstring GenerateRegistryKey(const std::wstring &displayName)
    {
        // Use timestamp for uniqueness, avoid auto-sorting
        static int counter = 0;
        counter++;

        wchar_t keyName[256];
        swprintf(keyName, 256, L"CustomApp_%s_%d", displayName.c_str(), counter);
        return keyName;
    }

    // Move registry item - implement actual order change by renaming registry keys
    bool MoveRegistryItem(int fromIndex, int toIndex)
    {
        if (fromIndex < 0 || toIndex < 0 || fromIndex >= (int)apps.size() || toIndex >= (int)apps.size())
            return false;

        AppEntry &fromApp = apps[fromIndex];
        AppEntry &toApp = apps[toIndex];

        // Only allow moving custom apps
        if (!fromApp.isCustom || !toApp.isCustom)
            return false;

        // Swap positions in list
        std::swap(apps[fromIndex], apps[toIndex]);

        // Update registry order - by renaming registry keys
        if (!UpdateRegistryOrder())
        {
            // If registry update fails, restore memory order
            std::swap(apps[fromIndex], apps[toIndex]);
            return false;
        }

        // Update list display
        UpdateListBoxDisplay();
        return true;
    }

    // Update registry order - recreate all custom items' registry keys to ensure correct order
    bool UpdateRegistryOrder()
    {
        // Collect all custom apps
        std::vector<AppEntry> customApps;
        for (const auto &app : apps)
        {
            if (app.isCustom)
            {
                customApps.push_back(app);
            }
        }

        // Recreate registry keys for each custom app (in new order)
        int position = 1;
        for (const auto &app : customApps)
        {
            // Generate new registry key name (with order number)
            wchar_t newKeyName[256];
            swprintf(newKeyName, 256, L"%02d_CustomApp_%s", position, app.displayName.c_str());

            // Skip if name hasn't changed
            if (app.name == newKeyName)
            {
                position++;
                continue;
            }

            // Create new registry key
            std::wstring newShellKey = L"Directory\\Background\\shell\\";
            newShellKey += newKeyName;

            HKEY hNewKey;
            if (RegCreateKeyExW(HKEY_CLASSES_ROOT, newShellKey.c_str(), 0, NULL, 0, KEY_ALL_ACCESS, NULL, &hNewKey, NULL) == ERROR_SUCCESS)
            {
                // Copy display name
                RegSetValueExW(hNewKey, NULL, 0, REG_SZ,
                               (const BYTE *)app.displayName.c_str(),
                               (app.displayName.length() + 1) * sizeof(wchar_t));

                // Copy icon
                if (!app.icon.empty())
                {
                    RegSetValueExW(hNewKey, L"Icon", 0, REG_SZ,
                                   (const BYTE *)app.icon.c_str(),
                                   (app.icon.length() + 1) * sizeof(wchar_t));
                }

                RegCloseKey(hNewKey);

                // Create command subkey
                std::wstring newCommandKey = newShellKey + L"\\command";
                if (RegCreateKeyExW(HKEY_CLASSES_ROOT, newCommandKey.c_str(), 0, NULL, 0, KEY_WRITE, NULL, &hNewKey, NULL) == ERROR_SUCCESS)
                {
                    std::wstring commandValue = L"\"";
                    commandValue += app.path;
                    commandValue += L"\"";

                    RegSetValueExW(hNewKey, NULL, 0, REG_SZ,
                                   (const BYTE *)commandValue.c_str(),
                                   (commandValue.length() + 1) * sizeof(wchar_t));
                    RegCloseKey(hNewKey);

                    // Delete old registry key
                    std::wstring oldShellKey = L"Directory\\Background\\shell\\";
                    oldShellKey += app.name;
                    DeleteRegistryTree(HKEY_CLASSES_ROOT, oldShellKey.c_str());
                }
                else
                {
                    // Failed to create command, delete shell key
                    RegDeleteKeyW(HKEY_CLASSES_ROOT, newShellKey.c_str());
                    return false;
                }
            }
            else
            {
                return false;
            }

            position++;
        }

        // Refresh system
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
        // Force reload from registry to ensure memory data matches registry
        ForceReloadFromRegistry();
        return true;
    }

    // Update list box display
    void UpdateListBoxDisplay()
    {
        SendMessageW(hListBox, LB_RESETCONTENT, 0, 0);
        for (const auto &app : apps)
        {
            std::wstring listText = GetDisplayText(app);
            SendMessageW(hListBox, LB_ADDSTRING, 0, (LPARAM)listText.c_str());
        }

        // Set horizontal scroll range
        UpdateHorizontalScroll();
    }

    // Update horizontal scroll range
    void UpdateHorizontalScroll()
    {
        if (!apps.empty())
        {
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

                // Calculate DPI scaling
                HDC hdcWindow = GetDC(hMainWindow);
                int dpiX = GetDeviceCaps(hdcWindow, LOGPIXELSX);
                ReleaseDC(hMainWindow, hdcWindow);
                float scale = dpiX / 96.0f;

                int horizontalExtent = maxWidth + (int)(100 * scale);
                SendMessage(hListBox, LB_SETHORIZONTALEXTENT, horizontalExtent, 0);
            }
        }
    }

    // Force reload all menu items from registry
    void ForceReloadFromRegistry()
    {
        // Clear existing data
        allApps.clear();
        apps.clear();

        // Reset list box
        SendMessageW(hListBox, LB_RESETCONTENT, 0, 0);

        // Reload directly from registry
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
                    // Modified custom app identification logic
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
        SortAppsByRegistryKeyName();
        FilterApps();
    }

public:
    RightClickManager() : hMainWindow(NULL), hListBox(NULL), hAddButton(NULL),
                          hRemoveButton(NULL), hRefreshButton(NULL), hShowAllCheckbox(NULL),
                          hMoveUpButton(NULL), hMoveDownButton(NULL),
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

    // Create context menu
    void CreateContextMenu()
    {
        hContextMenu = CreatePopupMenu();
        if (hContextMenu)
        {
            AppendMenuW(hContextMenu, MF_STRING, 1101, L"📁 Open in Registry");
            AppendMenuW(hContextMenu, MF_SEPARATOR, 0, NULL);
            AppendMenuW(hContextMenu, MF_STRING, 1102, L"🔄 Refresh This Item");
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
                // Calculate menu display position (below item)
                POINT pt;
                pt.x = itemRect.left + 10; // Slight offset to avoid covering text
                pt.y = itemRect.bottom;

                // Convert to screen coordinates
                ClientToScreen(hListBox, &pt);

                // Show right-click menu
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
        std::wstring regPath = L"Computer\\HKEY_CLASSES_ROOT\\Directory\\Background\\shell\\" + app.name;

        // Try to open Registry Editor using ShellExecute
        SHELLEXECUTEINFOW sei = {sizeof(sei)};
        sei.lpVerb = L"open";
        sei.lpFile = L"regedit.exe";
        sei.lpParameters = L""; // Don't use silent mode so user sees interface
        sei.nShow = SW_SHOW;

        if (ShellExecuteExW(&sei))
        {
            // Give Registry Editor time to start
            Sleep(1000);

            // Try to send keys to Registry Editor to navigate to our path
            HWND hRegEdit = FindWindowW(L"RegEdit_RegEdit", NULL);
            if (hRegEdit)
            {
                // Activate window
                SetForegroundWindow(hRegEdit);
                Sleep(100);

                // Send Ctrl+F to open search
                keybd_event(VK_CONTROL, 0, 0, 0);
                keybd_event('F', 0, 0, 0);
                keybd_event('F', 0, KEYEVENTF_KEYUP, 0);
                keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);

                // Wait for search dialog to open
                Sleep(500);

                // Enter our path in search box
                // First ensure input method is in English mode
                HWND hForeground = GetForegroundWindow();

                // Use more reliable way to input path
                for (wchar_t c : regPath)
                {
                    // For backslash, use virtual key code instead of character to avoid IME interference
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
                    Sleep(10); // Small delay to ensure correct character input
                }

                // Send Enter to start search
                Sleep(100);
                keybd_event(VK_RETURN, 0, 0, 0);
                keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0);

                std::wstring message = L"Registry Editor opened and attempted to locate specified position.\n"
                                       L"If not automatically located, please manually navigate to:\n" +
                                       regPath;

                MessageBoxW(hMainWindow,
                            message.c_str(),
                            L"Open Registry Location",
                            MB_OK | MB_ICONINFORMATION);
            }
            else
            {
                std::wstring message = L"Registry Editor opened but could not automatically locate.\n"
                                       L"Please manually navigate to:\n" +
                                       regPath;

                MessageBoxW(hMainWindow,
                            message.c_str(),
                            L"Open Registry Location",
                            MB_OK | MB_ICONINFORMATION);
            }
        }
        else
        {
            // Alternative method: show path for manual navigation
            std::wstring message = L"Please manually navigate to the following path in Registry Editor:\n\n" + regPath;

            MessageBoxW(hMainWindow,
                        message.c_str(),
                        L"Registry Location",
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
                        L"Can only rename items created by this program (✅ marked items)",
                        L"Information",
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

        // Adjust edit box position, avoid icon area
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

    // Finish editing - simplified version, remove auto-sorting
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
                    // Use RegCreateKeyExW with KEY_ALL_ACCESS permission
                    if (RegCreateKeyExW(HKEY_CLASSES_ROOT, shellKey.c_str(), 0, NULL,
                                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, NULL) == ERROR_SUCCESS)
                    {
                        RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE *)newName, (wcslen(newName) + 1) * sizeof(wchar_t));
                        RegCloseKey(hKey);

                        // Refresh system
                        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

                        // Re-read display name directly from registry for this item to ensure data sync
                        RefreshSingleItemFromRegistry(app.name);

                        MessageBoxW(hMainWindow, L"Rename successful!", L"Success", MB_OK | MB_ICONINFORMATION);
                    }
                    else
                    {
                        MessageBoxW(hMainWindow, L"Rename failed! Please run as administrator.", L"Error", MB_OK | MB_ICONERROR);
                    }
                }
                else
                {
                    // Name didn't change, no action needed
                    MessageBoxW(hMainWindow, L"Name unchanged.", L"Information", MB_OK | MB_ICONINFORMATION);
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
        HWND hExistingWindow = FindWindowW(L"RightClickManager", L"Desktop Context Menu Manager");
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
        // Create modern font
        NONCLIENTMETRICSW ncm = {sizeof(ncm)};
        SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
        hModernFont = CreateFontIndirectW(&ncm.lfMessageFont);

        // Initialize common controls
        INITCOMMONCONTROLSEX icex;
        icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
        icex.dwICC = ICC_LISTVIEW_CLASSES | ICC_TREEVIEW_CLASSES | ICC_BAR_CLASSES;
        InitCommonControlsEx(&icex);

        // Get DPI scaling factor
        HDC hdc = GetDC(NULL);
        int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
        ReleaseDC(NULL, hdc);
        float scale = dpiX / 96.0f;

        // Create main window
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

        // Calculate window size based on DPI scaling
        int windowWidth = (int)(800 * scale);  // Increased base width for English text
        int windowHeight = (int)(550 * scale); // Increased base height

        // Calculate window position to center it
        int screenWidth = GetSystemMetrics(SM_CXSCREEN);
        int screenHeight = GetSystemMetrics(SM_CYSCREEN);
        int x = (screenWidth - windowWidth) / 2;
        int y = (screenHeight - windowHeight) / 2;

        // Modify window style, remove resizable border
        hMainWindow = CreateWindowExW(
            WS_EX_APPWINDOW,
            L"RightClickManager",
            L"Desktop Context Menu Manager",
            WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
            x, y,
            windowWidth, windowHeight,
            NULL, NULL, hInstance, this);

        if (!hMainWindow)
            return false;

        // Apply font after main window creation
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
        // Get DPI scaling factor
        HDC hdc = GetDC(hMainWindow);
        int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
        ReleaseDC(hMainWindow, hdc);
        float scale = dpiX / 96.0f;

        // Adjust dimensions based on DPI scaling
        int listBoxWidth = (int)(600 * scale);  // Increased width for English text
        int listBoxHeight = (int)(500 * scale);
        int buttonWidth = (int)(160 * scale);   // Increased width for English text
        int buttonHeight = (int)(30 * scale);
        int margin = (int)(10 * scale);
        int rightPanelX = (int)(620 * scale);   // Adjusted for wider list box
        int helpTextWidth = (int)(170 * scale); // Increased width for English text
        int helpTextHeight = (int)(290 * scale);

        // List control - ensure includes vertical and horizontal scroll bars
        hListBox = CreateWindowExW(
            WS_EX_CLIENTEDGE,
            L"LISTBOX",
            L"",
            WS_CHILD | WS_VISIBLE | LBS_NOTIFY | LBS_HASSTRINGS |
                WS_VSCROLL | WS_HSCROLL | LBS_NOINTEGRALHEIGHT | LBS_DISABLENOSCROLL, // Add LBS_DISABLENOSCROLL to ensure scroll bars always available
            margin, margin,
            listBoxWidth, listBoxHeight,
            hMainWindow,
            (HMENU)1001,
            hInstance,
            NULL);

        // Set list item height for better readability
        int itemHeight = (int)(24 * scale); // Slightly increase item height for better readability
        SendMessage(hListBox, LB_SETITEMHEIGHT, 0, itemHeight);

        // Add button
        hAddButton = CreateWindowW(
            L"BUTTON",
            L"📁 Add Program",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            rightPanelX, margin,
            buttonWidth, buttonHeight,
            hMainWindow,
            (HMENU)1002,
            hInstance,
            NULL);

        // Remove button
        hRemoveButton = CreateWindowW(
            L"BUTTON",
            L"🗑️ Remove Selected",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            rightPanelX, margin + buttonHeight + margin / 2,
            buttonWidth, buttonHeight,
            hMainWindow,
            (HMENU)1003,
            hInstance,
            NULL);

        // Refresh button
        hRefreshButton = CreateWindowW(
            L"BUTTON",
            L"🔄 Refresh List",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            rightPanelX, margin + (buttonHeight + margin / 2) * 2,
            buttonWidth, buttonHeight,
            hMainWindow,
            (HMENU)1004,
            hInstance,
            NULL);

        // Move up button
        hMoveUpButton = CreateWindowW(
            L"BUTTON",
            L"⬆️ Move Up",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            rightPanelX, margin + (buttonHeight + margin / 2) * 3,
            buttonWidth, buttonHeight,
            hMainWindow,
            (HMENU)1006,
            hInstance,
            NULL);

        // Move down button
        hMoveDownButton = CreateWindowW(
            L"BUTTON",
            L"⬇️ Move Down",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            rightPanelX, margin + (buttonHeight + margin / 2) * 4,
            buttonWidth, buttonHeight,
            hMainWindow,
            (HMENU)1007,
            hInstance,
            NULL);

        // Create checkbox
        hShowAllCheckbox = CreateWindowW(
            L"BUTTON",
            L"Show All Items",
            WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
            rightPanelX, margin + (buttonHeight + margin / 2) * 5,
            buttonWidth, buttonHeight,
            hMainWindow,
            (HMENU)1005,
            hInstance,
            NULL);

        // Help text - updated to include move function description
        HWND hHelpText = CreateWindowW(
            L"STATIC",
            L"💡 Desktop Context Menu Management\n\n"
            L"✅ Items created by this program\n"
            L"📌 Items created by other programs\n\n"
            L"🖱️ Operation Tips:\n"
            L"• Double-click ✅ items to rename\n"
            L"• Right-click items for function menu\n"
            L"• Use ⬆️⬇️ buttons to adjust order\n"
            L"• Check box to show all items",
            WS_CHILD | WS_VISIBLE,
            rightPanelX, margin + (buttonHeight + margin / 2) * 6 + 10,
            helpTextWidth, helpTextHeight,
            hMainWindow,
            NULL,
            hInstance,
            NULL);

        // Apply modern font to all controls
        HWND hControls[] = {hListBox, hAddButton, hRemoveButton, hRefreshButton,
                            hMoveUpButton, hMoveDownButton, hShowAllCheckbox, hHelpText};
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
            // Preserve important information at beginning and end
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

        // Cancel editing if in progress
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
                            app.displayName = subkeyName; // Use registry key name if no display name
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
        SortAppsByRegistryKeyName();

        // Filter app list based on display settings
        FilterApps();
    }

    // Filter app list based on display settings
    void FilterApps()
    {
        // Cancel editing state if active
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
        UpdateHorizontalScroll();

        // Ensure vertical scroll bar displays correctly when too many items
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

        // If number of items exceeds visible area, vertical scroll bar will automatically display
        // We can select first item to give user visual feedback
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

        // Generate registry key name - simplified version, no auto-sorting
        std::wstring registryKey = GenerateRegistryKey(appName);

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
                swprintf(errorMsg, 256, L"Failed to set display name! Error code: %d", result);
                MessageBoxW(hMainWindow, errorMsg, L"Error", MB_OK | MB_ICONERROR);
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
            swprintf(errorMsg, 256, L"Failed to create registry key! Error code: %d", result);
            MessageBoxW(hMainWindow, errorMsg, L"Error", MB_OK | MB_ICONERROR);
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

                // Reload all menu items
                LoadAllContextMenuItems();
                return true;
            }
            else
            {
                // Add error information
                wchar_t errorMsg[256];
                swprintf(errorMsg, 256, L"Failed to set command! Error code: %d", result);
                MessageBoxW(hMainWindow, errorMsg, L"Error", MB_OK | MB_ICONERROR);
            }
        }
        else
        {
            // Add error information
            wchar_t errorMsg[256];
            swprintf(errorMsg, 256, L"Failed to create command subkey! Error code: %d", GetLastError());
            MessageBoxW(hMainWindow, errorMsg, L"Error", MB_OK | MB_ICONERROR);
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
            std::wstring warningMsg = L"Warning: This item was not created by this program and may be a system or other application's context menu item.\n\n";
            warningMsg += L"Name: " + app.displayName + L"\n";
            warningMsg += L"Path: " + app.path + L"\n\n";
            warningMsg += L"Are you sure you want to delete it?";

            if (MessageBoxW(hMainWindow, warningMsg.c_str(), L"Confirm System Item Deletion", MB_YESNO | MB_ICONWARNING) != IDYES)
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
                     L"Deletion failed! Error code: %d\n\n"
                     L"Possible reasons:\n"
                     L"• Registry key occupied by another process\n"
                     L"• Insufficient permissions\n"
                     L"• Registry key does not exist\n\n"
                     L"Please try running as administrator or restart and try again.",
                     errorCode);

            MessageBoxW(hMainWindow, errorMsg, L"Deletion Failed", MB_OK | MB_ICONERROR);
            return false;
        }

        // Refresh system
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);

        // Reload directly from registry
        ForceReloadFromRegistry();

        // Show success message
        MessageBoxW(hMainWindow,
                    L"Program removed from desktop context menu!\n"
                    L"If menu item still displays, try refreshing desktop (F5) or restarting Explorer.",
                    L"Deletion Successful", MB_OK | MB_ICONINFORMATION);

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

        // Note: index changes when deleting subkeys, so always enumerate from 0
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

    // Handle move up button click
    void OnMoveUpButtonClick()
    {
        int selectedIndex = (int)SendMessageW(hListBox, LB_GETCURSEL, 0, 0);
        if (selectedIndex == LB_ERR || selectedIndex <= 0)
        {
            MessageBoxW(hMainWindow, L"Please select a program first, and it cannot be the first item!", L"Information", MB_OK | MB_ICONINFORMATION);
            return;
        }

        // Only allow moving items created by this program
        if (!apps[selectedIndex].isCustom)
        {
            MessageBoxW(hMainWindow, L"Can only move items created by this program (✅ marked items)", L"Information", MB_OK | MB_ICONINFORMATION);
            return;
        }

        if (MoveRegistryItem(selectedIndex, selectedIndex - 1))
        {
            // Update selected item
            SendMessageW(hListBox, LB_SETCURSEL, selectedIndex - 1, 0);
            MessageBoxW(hMainWindow, L"Item moved up! Order in context menu also updated.", L"Success", MB_OK | MB_ICONINFORMATION);
        }
        else
        {
            MessageBoxW(hMainWindow, L"Move up failed! Please check if running as administrator or if move is legal.", L"Error", MB_OK | MB_ICONERROR);
        }
    }

    // Handle move down button click
    void OnMoveDownButtonClick()
    {
        int selectedIndex = (int)SendMessageW(hListBox, LB_GETCURSEL, 0, 0);
        if (selectedIndex == LB_ERR || selectedIndex >= (int)apps.size() - 1)
        {
            MessageBoxW(hMainWindow, L"Please select a program first, and it cannot be the last item!", L"Information", MB_OK | MB_ICONINFORMATION);
            return;
        }

        // Only allow moving items created by this program
        if (!apps[selectedIndex].isCustom)
        {
            MessageBoxW(hMainWindow, L"Can only move items created by this program (✅ marked items)", L"Information", MB_OK | MB_ICONINFORMATION);
            return;
        }

        if (MoveRegistryItem(selectedIndex, selectedIndex + 1))
        {
            // Update selected item
            SendMessageW(hListBox, LB_SETCURSEL, selectedIndex + 1, 0);
            MessageBoxW(hMainWindow, L"Item moved down! Order in context menu also updated.", L"Success", MB_OK | MB_ICONINFORMATION);
        }
        else
        {
            MessageBoxW(hMainWindow, L"Move down failed! Please check if running as administrator or if move is legal.", L"Error", MB_OK | MB_ICONERROR);
        }
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
        ofn.lpstrFilter = L"Executable Files\0*.exe\0All Files\0*.*\0";
        ofn.nFilterIndex = 1;
        ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

        if (GetOpenFileNameW(&ofn))
        {
            if (AddAppToContextMenu(fileName))
            {
                MessageBoxW(hMainWindow,
                            L"Program successfully added to desktop context menu!\n"
                            L"Program icon will also display in menu.\n"
                            L"May need to refresh desktop or restart Explorer to see changes.",
                            L"Success", MB_OK | MB_ICONINFORMATION);
            }
            else
            {
                MessageBoxW(hMainWindow, L"Failed to add program! Please run as administrator.", L"Error", MB_OK | MB_ICONERROR);
            }
        }
    }

    void OnRemoveButtonClick()
    {
        int selectedIndex = (int)SendMessageW(hListBox, LB_GETCURSEL, 0, 0);
        if (selectedIndex == LB_ERR)
        {
            MessageBoxW(hMainWindow, L"Please select a program first!", L"Information", MB_OK | MB_ICONINFORMATION);
            return;
        }

        AppEntry &app = apps[selectedIndex];
        std::wstring confirmMsg = L"Are you sure you want to remove this program from desktop context menu?\n\n";
        confirmMsg += L"Name: " + app.displayName + L"\n";
        confirmMsg += L"Path: " + app.path;

        if (MessageBoxW(hMainWindow, confirmMsg.c_str(), L"Confirm Deletion", MB_YESNO | MB_ICONQUESTION) == IDYES)
        {
            if (RemoveAppFromContextMenu(selectedIndex))
            {
                MessageBoxW(hMainWindow, L"Program removed from desktop context menu!", L"Success", MB_OK | MB_ICONINFORMATION);
            }
            else
            {
                MessageBoxW(hMainWindow, L"Failed to remove program!", L"Error", MB_OK | MB_ICONERROR);
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
        MessageBoxW(hMainWindow, L"Reloading menu items from registry...", L"Refreshing", MB_OK | MB_ICONINFORMATION);

        // Force reload all menu items from registry
        ForceReloadFromRegistry();

        // Show result statistics
        wchar_t resultMsg[256];
        swprintf(resultMsg, 256, L"Refresh complete!\nFound %d total menu items\n%d created by this program",
                 (int)allApps.size(),
                 (int)std::count_if(allApps.begin(), allApps.end(), [](const AppEntry &app)
                                    { return app.isCustom; }));

        MessageBoxW(hMainWindow, resultMsg, L"Refresh Complete", MB_OK | MB_ICONINFORMATION);
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
            else if (LOWORD(wParam) == 1006)
            { // Move up button
                OnMoveUpButtonClick();
            }
            else if (LOWORD(wParam) == 1007)
            { // Move down button
                OnMoveDownButtonClick();
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
            { // Context menu: Refresh this item
                if (contextMenuIndex >= 0 && contextMenuIndex < (int)apps.size())
                {
                    RefreshSingleItemFromRegistry(apps[contextMenuIndex].name);
                    MessageBoxW(hMainWindow, L"Selected item refreshed!", L"Refresh", MB_OK | MB_ICONINFORMATION);
                }
            }
            break;

        case WM_SIZE:
            // Recalculate horizontal scroll range when window size changes
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
                UpdateHorizontalScroll();
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
            // Get DPI scaling factor
            HDC hdc = GetDC(hMainWindow);
            int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
            ReleaseDC(hMainWindow, hdc);
            float scale = dpiX / 96.0f;

            // Lock window size based on DPI scaling
            MINMAXINFO *mmi = (MINMAXINFO *)lParam;
            mmi->ptMinTrackSize.x = (int)(800 * scale);  // Increased for English text
            mmi->ptMinTrackSize.y = (int)(550 * scale);  // Increased for English text
            mmi->ptMaxTrackSize.x = (int)(800 * scale);  // Increased for English text
            mmi->ptMaxTrackSize.y = (int)(550 * scale);  // Increased for English text
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
        // Fallback: try using older SetProcessDPIAware (supported on Vista and above)
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
        MessageBoxW(NULL,
                    L"This program requires administrator privileges to modify registry.\nPlease run as administrator.",
                    L"Insufficient Privileges",
                    MB_OK | MB_ICONWARNING);
        return 1;
    }

    RightClickManager manager;

    // Check if another instance is already running
    if (manager.IsAlreadyRunning())
    {
        manager.ActivateExistingInstance();
        return 0; // Exit directly, don't create new instance
    }

    if (!manager.Initialize(hInstance))
    {
        MessageBoxW(NULL, L"Program initialization failed!", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    return manager.Run();
}