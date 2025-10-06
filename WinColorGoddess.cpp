#include <iostream>
#include <string>
#include <ctime>
#include <windows.h>
#include <shlobj.h>
#include <wchar.h>
#include <shobjidl.h> // IDesktopWallpaper, IShellItemArray
#include <objbase.h>  // CoInitializeEx, CoUninitialize
#include <random> // 包含rand()和srand()

// 函数声明
void Color_TALOOC_Theme(bool darkMode);
bool Color_TALOOC_Dark();
void AddToStartup();
bool IsInStartup();
bool Color_TALOOC_SetWallpaperSlide(const wchar_t* path);
void Color_TALOOC_Wallpaper();


const std::wstring Color_TALOOC_NAME = L"WinColorGoddess";



bool Color_TALOOC_RanBool() {
    std::random_device rd;  // 获取随机种子
    std::mt19937 gen(rd()); // 初始化Mersenne Twister生成器
    std::bernoulli_distribution dist(0.5); // 伯努利分布
    return dist(gen);
}


int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    // 检查是否已经运行
    HANDLE hMutex = CreateMutex(NULL, TRUE, Color_TALOOC_NAME.c_str());
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        return 0; // 已有实例在运行
    }

    // 确保程序在开机启动项中
    if (!IsInStartup()) {
        AddToStartup();
    }

    // 主循环
    while (true) {
        bool useDarkTheme = Color_TALOOC_Dark();
        Color_TALOOC_Theme(useDarkTheme);
        Color_TALOOC_Wallpaper();

        // 0.1 min检查一次
        Sleep(1 * 60 * 1000);
    }

    ReleaseMutex(hMutex);
    return 0;
}


// 设置系统主题
void Color_TALOOC_Theme(bool darkMode) {
    HKEY hKey;
    DWORD value;
    value = darkMode ? 0 : 1; // 0=深色, 1=浅色
    if (RegOpenKeyEx(HKEY_CURRENT_USER,L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",0, KEY_WRITE, &hKey) == ERROR_SUCCESS) {
        // 设置Windows模式
        RegSetValueEx(hKey, L"SystemUsesLightTheme", 0, REG_DWORD, (const BYTE*)&value, sizeof(value));
        // 设置应用模式
        RegSetValueEx(hKey, L"AppsUseLightTheme", 0, REG_DWORD, (const BYTE*)&value, sizeof(value));
        DWORD customValue = 0; // 0=使用统一主题, 1=自定义模式
        RegSetValueEx(hKey, L"ColorPrevalence", 0, REG_DWORD, (const BYTE*)&customValue, sizeof(customValue));
        RegCloseKey(hKey);
    }

    //设置DWM-标题栏
    /*
    value = darkMode ? 1 : 0;
    if (RegOpenKeyEx(HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\DWM",
        0, KEY_WRITE, &hKey) == ERROR_SUCCESS) {
        RegSetValueEx(hKey, L"ColorPrevalence", 0, REG_DWORD, (const BYTE*)&value, sizeof(value));
        // 设置深色主题时启用彩色标题栏
        //if (darkMode) {
            //DWORD enableColor = 1;
            //RegSetValueEx(hKey, L"EnableWindowColorization", 0, REG_DWORD, (const BYTE*)&enableColor, sizeof(enableColor));
        //}
        RegCloseKey(hKey);
    }
    */

    //通知系统主题更改
    SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0,(LPARAM)L"ImmersiveColorSet", SMTO_ABORTIFHUNG, 1000, NULL);
}

bool Color_TALOOC_Dark() {
    SYSTEMTIME st;
    GetLocalTime(&st);
    // 18:00到6:00使用深色主题
    if (st.wHour >= 18 || st.wHour < 6) {
        return true;
    }
    return false;
    //return Color_TALOOC_RanBool();
}

// 根据时间切换壁纸幻灯片目录
void Color_TALOOC_Wallpaper() {
    bool isNight = Color_TALOOC_Dark();
    if (isNight) {
        // 夜间壁纸目录
        Color_TALOOC_SetWallpaperSlide(L"E:\\媒体\\壁纸\\-night");
    }
    else {
        // 日间壁纸目录
        Color_TALOOC_SetWallpaperSlide(L"E:\\媒体\\壁纸\\-day");
    }
}


/**
 * @brief 使用 IDesktopWallpaper 接口设置壁纸幻灯片目录
 * * @param path 包含图片的文件夹的宽字符路径
 * @return true 成功，false 失败
 */
bool Color_TALOOC_SetWallpaperSlide(const wchar_t* path) {
    //初始化 COM
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) return false;
    IDesktopWallpaper* pDesktopWallpaper = nullptr;
    IShellItem* pShellItem = nullptr;
    IShellItemArray* pShellItemArray = nullptr;
    //创建 IDesktopWallpaper 对象
    hr = CoCreateInstance(CLSID_DesktopWallpaper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pDesktopWallpaper));
    if (SUCCEEDED(hr)) {
        hr = SHCreateItemFromParsingName(
            path,NULL,
            IID_PPV_ARGS(&pShellItem)
        );
    }
    if (SUCCEEDED(hr)) {
        hr = SHCreateShellItemArrayFromShellItem(
            pShellItem,
            IID_PPV_ARGS(&pShellItemArray)
        );
    }
    if (SUCCEEDED(hr)) {
        //设置幻灯片目录
        hr = pDesktopWallpaper->SetSlideshow(pShellItemArray);
        //设置幻灯片选项，如打乱顺序 (DWSPO_SHUFFLE)
        // hr = pDesktopWallpaper->SetSlideshowOptions(DWSPO_SHUFFLE, 30000); 
    }
    //清理 COM 对象
    if (pDesktopWallpaper) pDesktopWallpaper->Release();
    if (pShellItem) pShellItem->Release();
    if (pShellItemArray) pShellItemArray->Release();
    //清理 COM
    CoUninitialize();
    return SUCCEEDED(hr);
}


//添加到开机启动
void AddToStartup() {
    HKEY hKey;
    WCHAR path[MAX_PATH];
    GetModuleFileName(NULL, path, MAX_PATH);
    if (RegOpenKeyEx(HKEY_CURRENT_USER,L"Software\\Microsoft\\Windows\\CurrentVersion\\Run",
        0, KEY_WRITE, &hKey) == ERROR_SUCCESS) {
        RegSetValueEx(hKey, Color_TALOOC_NAME.c_str(), 0, REG_SZ, (const BYTE*)path, (wcslen(path) + 1) * sizeof(WCHAR));
        RegCloseKey(hKey);
    }
}

// 检查是否在开机启动项中
bool IsInStartup() {
    HKEY hKey;
    WCHAR path[MAX_PATH];
    DWORD size = MAX_PATH;
    if (RegOpenKeyEx(HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Run",
        0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        if (RegQueryValueEx(hKey, Color_TALOOC_NAME.c_str(), NULL, NULL, (LPBYTE)path, &size) == ERROR_SUCCESS) {
            RegCloseKey(hKey);
            WCHAR currentPath[MAX_PATH];
            GetModuleFileName(NULL, currentPath, MAX_PATH);
            return wcscmp(path, currentPath) == 0;
        }
        RegCloseKey(hKey);
    }
    return false;
}
