#define  _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <string>
#include <vector>
#include <sstream>
#include <objbase.h>
#include <string>
#include <sstream>
#include <vector>
#include <utility>
#include <Windows.h>
#include <comdef.h>  // For _bstr_t (optional for debug)
#include <windows.h>
#include <vector>
#include <fstream>
#include <iostream>
#include <windows.h>
#include <commdlg.h>
#include <tchar.h>
#include <Windows.h>
#include <ShlObj.h>
#include <Windows.h>
#include <ShlObj.h>
#include <Objbase.h>
#include <atlimage.h> // For handling CImage
#include <gdiplus.h>  // For GDI+ functionality
#include <atlimage.h> // For saving the icon as a BMP/PNG
#include <string>

#include <windows.h>
#include <shobjidl_core.h>
#include <gdiplus.h>
#include <fstream>
#include <vector>
#include <windows.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <string>
#include <vector>
#include <fstream>

#include <windows.h>
#include <gdiplus.h>
#include <vector>
#include <wincodec.h>
#include <vector>
#include <comdef.h> 
using namespace Gdiplus;

#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "User32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Windowscodecs.lib")



// Helper to resolve .lnk file
std::string ResolveShortcut(const std::string& lnkPath) {
	std::string resolvedPath;
	CoInitialize(nullptr); // Initialize COM library

	IShellLinkW* pShellLink = nullptr;
	if (SUCCEEDED(CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_IShellLinkW, (void**)&pShellLink))) {
		IPersistFile* pPersistFile = nullptr;
		if (SUCCEEDED(pShellLink->QueryInterface(IID_IPersistFile, (void**)&pPersistFile))) {
			wchar_t wsz[MAX_PATH];
			MultiByteToWideChar(CP_ACP, 0, lnkPath.c_str(), -1, wsz, MAX_PATH);
			if (SUCCEEDED(pPersistFile->Load(wsz, STGM_READ))) {
				wchar_t targetPath[MAX_PATH];
				if (SUCCEEDED(pShellLink->GetPath(targetPath, MAX_PATH, nullptr, 0))) {
					// Convert wide character path to multi-byte (narrow character) path
					char targetPathA[MAX_PATH];
					WideCharToMultiByte(CP_ACP, 0, targetPath, -1, targetPathA, MAX_PATH, nullptr, nullptr);
					resolvedPath = targetPathA;

					// Replace backslashes with double backslashes for JSON compatibility
					size_t pos = 0;
					while ((pos = resolvedPath.find("\\", pos)) != std::string::npos) {
						resolvedPath.replace(pos, 1, "\\\\");
						pos += 2; // Skip past the replacement
					}
				}
			}
			pPersistFile->Release();
		}
		pShellLink->Release();
	}
	CoUninitialize(); // Uninitialize COM library
	return resolvedPath;
}

// Helper to scan directories
void ScanDirectoryA(const std::string& path, std::vector<std::pair<std::string, std::string>>& results) {
	WIN32_FIND_DATAA findData;
	HANDLE hFind = FindFirstFileA((path + "\\*.lnk").c_str(), &findData);
	if (hFind != INVALID_HANDLE_VALUE) {
		do {
			if (!(findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
				std::string lnkPath = path + "\\" + findData.cFileName;
				std::string resolvedPath = ResolveShortcut(lnkPath);
				if (!resolvedPath.empty()) {
					// Remove the ".lnk" suffix from the name
					std::string name = findData.cFileName;
					if (name.size() > 4 && name.substr(name.size() - 4) == ".lnk") {
						name = name.substr(0, name.size() - 4); // Remove the ".lnk" part
					}
					results.emplace_back(name, resolvedPath);
				}
			}
		} while (FindNextFileA(hFind, &findData));
		FindClose(hFind);
	}
}


// 解析快捷方式 .lnk 文件，获取目标路径
bool ResolveShortcutW(const std::wstring& lnkPath, std::wstring& targetPath) {
	HRESULT hr;
	IShellLinkW* pShellLink = nullptr;
	WIN32_FIND_DATAW findData;

	// 初始化 COM 库
	hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
	if (FAILED(hr)) {
		return false;
	}

	// 创建 IShellLink 对象
	hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_IShellLinkW, (void**)&pShellLink);
	if (FAILED(hr)) {
		CoUninitialize();
		return false;
	}

	// 创建一个 IPersistFile 接口用于加载 .lnk 文件
	IPersistFile* pPersistFile = nullptr;
	hr = pShellLink->QueryInterface(IID_IPersistFile, (void**)&pPersistFile);
	if (FAILED(hr)) {
		pShellLink->Release();
		CoUninitialize();
		return false;
	}

	// 加载 .lnk 文件
	hr = pPersistFile->Load(lnkPath.c_str(), STGM_READ);
	pPersistFile->Release();
	if (FAILED(hr)) {
		pShellLink->Release();
		CoUninitialize();
		return false;
	}

	// 获取快捷方式的目标路径
	TCHAR szTargetPath[MAX_PATH] = { 0 };
	hr = pShellLink->GetPath(szTargetPath, MAX_PATH, &findData, SLGP_SHORTPATH);
	if (FAILED(hr)) {
		pShellLink->Release();
		CoUninitialize();
		return false;
	}

	// 设置目标路径
	targetPath = szTargetPath;

	// 释放资源
	pShellLink->Release();
	CoUninitialize();

	return true;
}

void ScanDirectory(const std::wstring& directory, std::vector<std::pair<std::wstring, std::wstring>>& apps) {
	WIN32_FIND_DATAW findData;
	HANDLE hFind = FindFirstFileW((directory + L"\\*.lnk").c_str(), &findData);

	if (hFind != INVALID_HANDLE_VALUE) {
		do {
			const std::wstring fileName = findData.cFileName;
			const std::wstring fullPath = directory + L"\\" + findData.cFileName;

			// 跳过特殊目录 "." 和 ".."
			if (fileName == L"." || fileName == L"..") continue;

			// 如果是文件，且是快捷方式，解析目标路径
			if (!(findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
				std::wstring targetPath;
				if (ResolveShortcutW(fullPath, targetPath)) {

					std::wstring name = fileName;
					// 如果文件名以 .lnk 结尾，去掉 .lnk 后缀
					if (name.size() > 4 && name.substr(name.size() - 4) == L".lnk") {
						name = name.substr(0, name.size() - 4); // Remove the ".lnk" part
					}

					apps.emplace_back(name, targetPath);
				}
			}
			// 如果是子目录，递归扫描
			else if (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
				ScanDirectory(fullPath, apps);
			}

		} while (FindNextFileW(hFind, &findData) != 0);

		FindClose(hFind);
	}
}

extern "C" __declspec(dllexport) HICON GetFileIcon1(const char* extension)
{
	HICON icon = NULL;
	if (extension && strlen(extension) > 0)
	{
		SHFILEINFOA info;
		if (SHGetFileInfoA(extension,
			FILE_ATTRIBUTE_NORMAL,
			&info,
			sizeof(info),
			SHGFI_SYSICONINDEX | SHGFI_ICON | SHGFI_USEFILEATTRIBUTES))
		{
			icon = info.hIcon;
		}
	}
	return icon;
}


extern "C" __declspec(dllexport) const wchar_t* GetStartMenuApps() {
	static std::wstring result;
	std::vector<std::pair<std::wstring, std::wstring>> apps;

	// 获取环境变量路径，使用 wchar_t*
	wchar_t* userPath = _wgetenv(L"APPDATA");
	wchar_t* systemPath = _wgetenv(L"PROGRAMDATA");

	// 组合完整路径
	std::wstring userStartMenu = std::wstring(userPath) + L"\\Microsoft\\Windows\\Start Menu\\Programs";
	std::wstring systemStartMenu = std::wstring(systemPath) + L"\\Microsoft\\Windows\\Start Menu\\Programs";

	// 扫描目录
	ScanDirectory(userStartMenu, apps);
	ScanDirectory(systemStartMenu, apps);

	// 格式化为 JSON 样式的字符串
	std::wostringstream oss;
	oss << L"[";

	for (size_t i = 0; i < apps.size(); ++i) {
		// 转换路径中的单反斜杠为双反斜杠
		std::wstring path = apps[i].second;
		for (size_t j = 0; j < path.length(); ++j) {
			if (path[j] == L'\\') {
				path.insert(j, L"\\");  // 在每个单反斜杠之前插入一个反斜杠
				j++;  // 跳过下一个字符，因为我们已经插入了一个新的反斜杠
			}
		}

		// 生成 JSON 项
		oss << L"{\"name\":\"" << apps[i].first << L"\",\"path\":\"" << path << L"\"}";

		if (i < apps.size() - 1) oss << L",";  // 如果不是最后一个元素，加逗号
	}
	oss << L"]";

	result = oss.str();

	// 返回宽字符字符串
	return result.c_str();
}



