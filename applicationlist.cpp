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


// ������ݷ�ʽ .lnk �ļ�����ȡĿ��·��
bool ResolveShortcutW(const std::wstring& lnkPath, std::wstring& targetPath) {
	HRESULT hr;
	IShellLinkW* pShellLink = nullptr;
	WIN32_FIND_DATAW findData;

	// ��ʼ�� COM ��
	hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
	if (FAILED(hr)) {
		return false;
	}

	// ���� IShellLink ����
	hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_IShellLinkW, (void**)&pShellLink);
	if (FAILED(hr)) {
		CoUninitialize();
		return false;
	}

	// ����һ�� IPersistFile �ӿ����ڼ��� .lnk �ļ�
	IPersistFile* pPersistFile = nullptr;
	hr = pShellLink->QueryInterface(IID_IPersistFile, (void**)&pPersistFile);
	if (FAILED(hr)) {
		pShellLink->Release();
		CoUninitialize();
		return false;
	}

	// ���� .lnk �ļ�
	hr = pPersistFile->Load(lnkPath.c_str(), STGM_READ);
	pPersistFile->Release();
	if (FAILED(hr)) {
		pShellLink->Release();
		CoUninitialize();
		return false;
	}

	// ��ȡ��ݷ�ʽ��Ŀ��·��
	TCHAR szTargetPath[MAX_PATH] = { 0 };
	hr = pShellLink->GetPath(szTargetPath, MAX_PATH, &findData, SLGP_SHORTPATH);
	if (FAILED(hr)) {
		pShellLink->Release();
		CoUninitialize();
		return false;
	}

	// ����Ŀ��·��
	targetPath = szTargetPath;

	// �ͷ���Դ
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

			// ��������Ŀ¼ "." �� ".."
			if (fileName == L"." || fileName == L"..") continue;

			// ������ļ������ǿ�ݷ�ʽ������Ŀ��·��
			if (!(findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
				std::wstring targetPath;
				if (ResolveShortcutW(fullPath, targetPath)) {

					std::wstring name = fileName;
					// ����ļ����� .lnk ��β��ȥ�� .lnk ��׺
					if (name.size() > 4 && name.substr(name.size() - 4) == L".lnk") {
						name = name.substr(0, name.size() - 4); // Remove the ".lnk" part
					}

					apps.emplace_back(name, targetPath);
				}
			}
			// �������Ŀ¼���ݹ�ɨ��
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

	// ��ȡ��������·����ʹ�� wchar_t*
	wchar_t* userPath = _wgetenv(L"APPDATA");
	wchar_t* systemPath = _wgetenv(L"PROGRAMDATA");

	// �������·��
	std::wstring userStartMenu = std::wstring(userPath) + L"\\Microsoft\\Windows\\Start Menu\\Programs";
	std::wstring systemStartMenu = std::wstring(systemPath) + L"\\Microsoft\\Windows\\Start Menu\\Programs";

	// ɨ��Ŀ¼
	ScanDirectory(userStartMenu, apps);
	ScanDirectory(systemStartMenu, apps);

	// ��ʽ��Ϊ JSON ��ʽ���ַ���
	std::wostringstream oss;
	oss << L"[";

	for (size_t i = 0; i < apps.size(); ++i) {
		// ת��·���еĵ���б��Ϊ˫��б��
		std::wstring path = apps[i].second;
		for (size_t j = 0; j < path.length(); ++j) {
			if (path[j] == L'\\') {
				path.insert(j, L"\\");  // ��ÿ������б��֮ǰ����һ����б��
				j++;  // ������һ���ַ�����Ϊ�����Ѿ�������һ���µķ�б��
			}
		}

		// ���� JSON ��
		oss << L"{\"name\":\"" << apps[i].first << L"\",\"path\":\"" << path << L"\"}";

		if (i < apps.size() - 1) oss << L",";  // ����������һ��Ԫ�أ��Ӷ���
	}
	oss << L"]";

	result = oss.str();

	// ���ؿ��ַ��ַ���
	return result.c_str();
}



