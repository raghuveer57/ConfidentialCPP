// Confidential.cpp : Defines the entry point for the application.
//

#include "framework.h"
#include "Confidential.h"
#include <shellapi.h>
#include <string>
#include <unordered_set>
#include <unordered_map>
#include <Psapi.h>
#include <cstdlib>
#include <propvarutil.h>
#include <propkey.h>
#include <shobjidl.h>
#include <propsys.h>
#include <strsafe.h>
#include <comdef.h>
#include <Wbemidl.h>
#include <array>
#include <iostream>
#include <string>
#include <windows.h>
#include <sstream>
#include <fstream>

#pragma comment(lib, "wbemuuid.lib")

#define MAX_LOADSTRING 100
#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 1001
#define TIMER_ID 1

// Global Variables:
HINSTANCE hInst;                                // current instance
HWND g_hWnd;
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name
NOTIFYICONDATA g_notifyIconData;  
std::unordered_set<std::wstring> g_targetProcesses = {
    L"WINWORD.EXE", L"EXCEL.EXE", L"POWERPNT.EXE" /*L"ONENOTE.EXE", L"Acrobat.exe", L"ACRORD32.EXE"*/
};
std::unordered_map<HWND, HWND> g_buttonMap; // Map to keep track of buttons for each target window
std::unordered_map<HWND, HWND> g_panMap;
std::unordered_map<HWND, HWND> g_targetMap;
std::unordered_map<HWND, bool> g_checkMap;
std::unordered_map<HWND, std::wstring> g_pathMap;
HBRUSH hBrush;

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK ButtonWndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK PaneWndProc(HWND, UINT, WPARAM, LPARAM);
void CreateTrayIcon(HWND);
void RemoveTrayIcon();
void ShowContextMenu(HWND, POINT);
void UpdateButtonPosition();
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam);
bool IsTargetProcess(const std::wstring& processName);
std::wstring GetProcessName(DWORD processID);
void CreateOrUpdateButton(HWND targetWnd);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{

    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: Place code here.

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_CONFIDENTIAL, szWindowClass, MAX_LOADSTRING);
    if (FindWindow(szWindowClass, szTitle)) {
      MessageBox(NULL, L"This app is already running", L"Error", MB_OKCANCEL);
      return 0;
    }

    MyRegisterClass(hInstance);

    // Register the button window class
    WNDCLASSW wcButton = {};
    wcButton.lpfnWndProc = ButtonWndProc;
    wcButton.hInstance = hInstance;
    wcButton.lpszClassName = _T("ButtonWindowClass");
    wcButton.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wcButton.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClassW(&wcButton);

    WNDCLASSW wcPane = {};
    wcPane.lpfnWndProc = PaneWndProc;
    wcPane.hInstance = hInstance;
    wcPane.lpszClassName = _T("PaneWindowClass");
    wcPane.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    RegisterClassW(&wcPane);

    // Perform application initialization:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_CONFIDENTIAL));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    // Clean up
    //RemoveTrayIcon();
    KillTimer(g_hWnd, TIMER_ID);

    return (int) msg.wParam;
}


//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSW wcex = {};
    wcex.lpfnWndProc    = WndProc;
    wcex.hInstance      = hInstance;
    wcex.lpszClassName  = szWindowClass;

    return RegisterClassW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowExW(WS_EX_TOOLWINDOW, szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   g_hWnd = hWnd;

   if (!hWnd)
   {
      return FALSE;
   }

   // Create the tray icon
   //CreateTrayIcon(hWnd);

   // Set timer for monitoring processes
   SetTimer(hWnd, TIMER_ID, 10, NULL); // 0.01-second interval

   return TRUE;
}

HANDLE OpenFileForReadWrite(const std::wstring& filePath) {
  HANDLE hFile = CreateFileW(
    filePath.c_str(),
    GENERIC_READ | GENERIC_WRITE,
    FILE_SHARE_READ,
    NULL,
    OPEN_EXISTING,
    FILE_ATTRIBUTE_NORMAL,
    NULL
  );
  if (hFile == INVALID_HANDLE_VALUE) {
    
  }
  return hFile;
}

std::string ReadContent(HANDLE hFile) {
  if (hFile == INVALID_HANDLE_VALUE) {
    return "";
  }

  DWORD fileSize = GetFileSize(hFile, NULL);
  if (fileSize == INVALID_FILE_SIZE) {
    
    return "";
  }

  char* buffer = new char[fileSize + 1];
  DWORD bytesRead;
  SetFilePointer(hFile, 0, NULL, FILE_BEGIN); // Move pointer to the beginning of the file
  if (!ReadFile(hFile, buffer, fileSize, &bytesRead, NULL)) {
    
    delete[] buffer;
    return "";
  }

  buffer[bytesRead] = '\0';
  std::string content(buffer);
  delete[] buffer;

  return content;
}

bool WriteContent(HANDLE hFile, const std::string& content) {
  if (hFile == INVALID_HANDLE_VALUE) {
    return false;
  }

  SetFilePointer(hFile, 0, NULL, FILE_BEGIN); // Move pointer to the beginning of the file
  DWORD bytesWritten;
  if (!WriteFile(hFile, content.c_str(), content.size(), &bytesWritten, NULL)) {
    
    return false;
  }

  // Optionally, truncate the file if the new content is smaller
  SetEndOfFile(hFile);

  return true;
}


// Implementation of the functions
std::string ConvertWideToNarrow(const std::wstring& wideString) {
    int bufferSize = WideCharToMultiByte(CP_UTF8, 0, wideString.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string narrowString(bufferSize, 0);
    WideCharToMultiByte(CP_UTF8, 0, wideString.c_str(), -1, &narrowString[0], bufferSize, nullptr, nullptr);
    return narrowString;
}

std::string RunCommand(const std::wstring& command) {

    //return result;
    STARTUPINFO startupInfo;
    PROCESS_INFORMATION processInfo;
    SECURITY_ATTRIBUTES securityAttributes;
    HANDLE readPipe, writePipe;

    ZeroMemory(&startupInfo, sizeof(startupInfo));
    ZeroMemory(&processInfo, sizeof(processInfo));
    ZeroMemory(&securityAttributes, sizeof(securityAttributes));

    startupInfo.cb = sizeof(startupInfo);
    startupInfo.dwFlags |= STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    startupInfo.wShowWindow = SW_HIDE;

    securityAttributes.nLength = sizeof(SECURITY_ATTRIBUTES);
    securityAttributes.bInheritHandle = TRUE;
    securityAttributes.lpSecurityDescriptor = NULL;

    if (!CreatePipe(&readPipe, &writePipe, &securityAttributes, 0)) {
        throw std::runtime_error("Failed to create pipe");
    }

    startupInfo.hStdOutput = writePipe;
    startupInfo.hStdError = writePipe;

    std::wstring commandLine = L"cmd.exe /C " + command;

    if (!CreateProcess(
        NULL,
        &commandLine[0],
        NULL,
        NULL,
        TRUE,
        CREATE_NO_WINDOW,
        NULL,
        NULL,
        &startupInfo,
        &processInfo
    )) {
        CloseHandle(readPipe);
        CloseHandle(writePipe);
        throw std::runtime_error("Failed to create process");
    }

    CloseHandle(writePipe);

    std::array<char, 128> buffer;
    std::string result;
    DWORD bytesRead;

    while (ReadFile(readPipe, buffer.data(), buffer.size(), &bytesRead, NULL) && bytesRead > 0) {
        result.append(buffer.data(), bytesRead);
    }

    CloseHandle(readPipe);
    CloseHandle(processInfo.hProcess);
    CloseHandle(processInfo.hThread);

    return result;
}

std::string RunPowerShellCommand(const std::wstring& command) {
  // Prepare the pipes for stdout redirection
  SECURITY_ATTRIBUTES sa;
  sa.nLength = sizeof(SECURITY_ATTRIBUTES);
  sa.bInheritHandle = TRUE;
  sa.lpSecurityDescriptor = NULL;

  HANDLE hReadPipe, hWritePipe;
  if (!CreatePipe(&hReadPipe, &hWritePipe, &sa, 0)) {
    return "";
  }

  // Ensure the read handle to the pipe is not inherited
  SetHandleInformation(hReadPipe, HANDLE_FLAG_INHERIT, 0);

  // Prepare the process startup information
  STARTUPINFOW si;
  ZeroMemory(&si, sizeof(si));
  si.cb = sizeof(si);
  si.dwFlags |= STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
  si.wShowWindow = SW_HIDE;
  si.hStdOutput = hWritePipe;
  si.hStdError = hWritePipe;

  // Prepare the process information
  PROCESS_INFORMATION pi;
  ZeroMemory(&pi, sizeof(pi));

  // Build the command line for PowerShell
  std::wstring commandLine = L"powershell.exe -NoProfile -Command \"" + command + L"\"";

  // Create the process
  if (!CreateProcessW(NULL, &commandLine[0], NULL, NULL, TRUE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi)) {
    CloseHandle(hWritePipe);
    CloseHandle(hReadPipe);
    return "";
  }

  // Close the write end of the pipe since the child process has it now
  CloseHandle(hWritePipe);

  // Wait for the process to complete
  WaitForSingleObject(pi.hProcess, INFINITE);

  // Check the exit code
  DWORD exitCode;
  if (!GetExitCodeProcess(pi.hProcess, &exitCode) || exitCode != 0) {
    CloseHandle(hReadPipe);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);
    return "";
  }

  // Read the output from the child process
  const DWORD bufferSize = 4096;
  char buffer[bufferSize];
  DWORD bytesRead;
  std::string output;

  while (ReadFile(hReadPipe, buffer, bufferSize - 1, &bytesRead, NULL) && bytesRead > 0) {
    buffer[bytesRead] = '\0';
    output += buffer;
  }

  // Cleanup
  CloseHandle(hReadPipe);
  CloseHandle(pi.hProcess);
  CloseHandle(pi.hThread);

  return output;
}

std::wstring GetCommandLineFromProcess(int processId)
{
  HRESULT hres;

  // Initialize COM
  hres = CoInitializeEx(0, COINIT_MULTITHREADED);
  if (FAILED(hres))
  {
    return L"";
  }

  // Initialize security
  hres = CoInitializeSecurity(
    NULL,
    -1,                          // COM negotiates service
    NULL,                        // Authentication services
    NULL,                        // Reserved
    RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication 
    RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation  
    NULL,                        // Authentication info
    EOAC_NONE,                   // Additional capabilities 
    NULL                         // Reserved
  );

  if (FAILED(hres))
  {
    CoUninitialize();
    return L"";
  }

  // Obtain the initial locator to WMI
  IWbemLocator* pLoc = NULL;

  hres = CoCreateInstance(
    CLSID_WbemLocator,
    0,
    CLSCTX_INPROC_SERVER,
    IID_IWbemLocator, (LPVOID*)&pLoc);

  if (FAILED(hres))
  {
    CoUninitialize();
    return L"";
  }

  // Connect to WMI namespace
  IWbemServices* pSvc = NULL;

  hres = pLoc->ConnectServer(
    _bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
    NULL,                    // User name
    NULL,                    // User password
    0,                       // Locale 
    NULL,                    // Security flags
    0,                       // Authority 
    0,                       // Context object 
    &pSvc                    // IWbemServices proxy
  );

  if (FAILED(hres))
  {
    pLoc->Release();
    CoUninitialize();
    return L"";
  }

  // Set security levels on the proxy
  hres = CoSetProxyBlanket(
    pSvc,                        // Indicates the proxy to set
    RPC_C_AUTHN_WINNT,           // RPC_C_AUTHN_xxx
    RPC_C_AUTHZ_NONE,            // RPC_C_AUTHZ_xxx
    NULL,                        // Server principal name 
    RPC_C_AUTHN_LEVEL_CALL,      // RPC_C_AUTHN_LEVEL_xxx 
    RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
    NULL,                        // Client identity
    EOAC_NONE                    // Proxy capabilities 
  );

  if (FAILED(hres))
  {
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();
    return L"";
  }

  // Query for the command line of the specified process
  IEnumWbemClassObject* pEnumerator = NULL;
  std::wstring query = L"SELECT CommandLine FROM Win32_Process WHERE ProcessId = " + std::to_wstring(processId);

  hres = pSvc->ExecQuery(
    bstr_t("WQL"),
    bstr_t(query.c_str()),
    WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
    NULL,
    &pEnumerator);

  if (FAILED(hres))
  {
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();
    return L"";
  }

  // Get the data from the query
  IWbemClassObject* pclsObj = NULL;
  ULONG uReturn = 0;
  std::wstring commandLine;

  while (pEnumerator)
  {
    HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

    if (0 == uReturn)
    {
      break;
    }

    VARIANT vtProp;
    hr = pclsObj->Get(L"CommandLine", 0, &vtProp, 0, 0);

    if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR)
    {
      commandLine = vtProp.bstrVal;
    }

    VariantClear(&vtProp);
    pclsObj->Release();
  }

  // Cleanup
  pEnumerator->Release();
  pSvc->Release();
  pLoc->Release();
  CoUninitialize();

  return commandLine;
}

std::wstring GetSubstringFromLastQuote(const std::wstring& s)
{
  if (s.empty())
  {
    return L"";
  }

  size_t lastIndex = s.rfind(L':');
  if (lastIndex == std::wstring::npos)
  {
    return L""; // No colon found in the string
  }

  std::wstring result;
  for (size_t i = lastIndex - 1; i < s.length() && s[i] != L'\"'; ++i)
  {
    result += s[i];
  }

  return result;
}

std::wstring string_to_wstring(const std::string& str) {
  // Calculate the size of the destination buffer
  size_t requiredSize = 0;
  mbstowcs_s(&requiredSize, nullptr, 0, str.c_str(), _TRUNCATE);

  // Allocate the buffer
  std::vector<wchar_t> buffer(requiredSize);

  // Perform the conversion
  mbstowcs_s(&requiredSize, buffer.data(), buffer.size(), str.c_str(), _TRUNCATE);

  // Return the result as a std::wstring
  return std::wstring(buffer.data());
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE: Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_CREATE:
      hBrush = CreateSolidBrush(RGB(0, 255, 255)); // Red color
      break;
    case WM_TRAYICON:
      if (lParam == WM_RBUTTONUP) {
        POINT pt;
        GetCursorPos(&pt);
        ShowContextMenu(hWnd, pt);
      }
      break;
    case WM_COMMAND:
      if (LOWORD(wParam) == ID_TRAY_EXIT) {
        DestroyWindow(hWnd);
      }
      break;
    case WM_TIMER:
      if (wParam == TIMER_ID) {
        UpdateButtonPosition();
      }
      break;
    case WM_DESTROY:
        DeleteObject(hBrush);
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

HRGN CreateRoundedRegion(int width, int height) {
  // Define the rounded region
  HRGN hRgn = CreateRectRgn(0, 0, 0, 0); // Initial empty region

  HRGN hRectRgn = CreateRectRgn(0, 0, width - 25, height - 1);
  HRGN hRoundedRgn = CreateRoundRectRgn(0, 0, width, height, 5, 5);  // Top right corner

  CombineRgn(hRgn, hRectRgn, hRoundedRgn, RGN_OR);

  // Clean up
  DeleteObject(hRectRgn);
  DeleteObject(hRoundedRgn);

  return hRgn;
}

// Button window procedure
LRESULT CALLBACK ButtonWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  HWND targetWnd = g_targetMap[g_panMap[hwnd]], paneWnd = g_panMap[hwnd];
  RECT rect;
  GetWindowRect(targetWnd, &rect);
  
  switch (msg) {
  case WM_CREATE: {
    HRGN hRgn = CreateRoundedRegion(40, 32);
    SetWindowRgn(hwnd, hRgn, TRUE);
    SetTimer(hwnd, 10, 100, NULL);
  }
    break;
  case WM_LBUTTONDOWN:
    g_checkMap[hwnd] = true;
    ShowWindow(hwnd, SW_HIDE);
    break;
  case WM_PAINT: {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hwnd, &ps);
    // Create a compatible device context for double buffering
    HDC hdcMem = CreateCompatibleDC(hdc);
    HBITMAP hBitmap = CreateCompatibleBitmap(hdc, 100, 100);
    SelectObject(hdcMem, hBitmap);
    HBRUSH redBrush = CreateSolidBrush(RGB(56, 65, 74));
    RECT backRect = { 0, 0, 40, 32 };
    FillRect(hdcMem, &backRect, redBrush);
    // Load the icon from a resource
    HICON hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_GEAR));
    DrawIconEx(hdcMem, 12, 4, hIcon, 22, 22, 0, NULL, DI_NORMAL);
    // Copy contents of the off-screen bitmap to the window's device context
    BitBlt(hdc, 0, 0, 40, 32, hdcMem, 0, 0, SRCCOPY);
    // Cleanup
    DeleteDC(hdcMem);
    DeleteObject(hBitmap);
    EndPaint(hwnd, &ps);
    break;
  }
  case WM_TIMER:
    if (!IsWindow(targetWnd)) {
      KillTimer(hwnd, 100);
      DestroyWindow(hwnd);
    }
    break;
  case WM_CLOSE:
    KillTimer(hwnd, 100);
    DestroyWindow(hwnd);
    break;
  default:
    return DefWindowProc(hwnd, msg, wParam, lParam);
  }
  return 0;
}

std::wstring GetExecutablePath()
{
    wchar_t buffer[MAX_PATH];
    GetModuleFileName(NULL, buffer, MAX_PATH);
    return std::wstring(buffer);
}

std::wstring GetExecutableDirectory()
{
    std::wstring exePath = GetExecutablePath();
    std::size_t pos = exePath.find_last_of(L"\\/");
    return exePath.substr(0, pos);
}

// Button window procedure
LRESULT CALLBACK PaneWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  HWND targetWnd = g_targetMap[hwnd], buttonWnd = g_buttonMap[g_targetMap[hwnd]];
  RECT rect;
  GetWindowRect(targetWnd, &rect);
  switch (msg) {
  case WM_COMMAND:
    if (LOWORD(wParam) == 1) {
      ShowWindow(hwnd, SW_HIDE);
      g_checkMap[buttonWnd] = false;
    }
    else if (LOWORD(wParam == 2)) {
        std::wstring exeDirectory = GetExecutableDirectory();
        std::wstring exeName = L"WordDocumentUpdater.exe";
        std::wstring exePath = exeDirectory + L"\\" + exeName;
        exePath = L"\"" + exePath + L"\"";

        std::wstring command = exePath + L" get";
        std::string output;

        try {
            output = RunCommand(command);
        }
        catch (const std::exception& e) {
            std::wstring errorMessage = L"Error running command: " + std::wstring(e.what(), e.what() + strlen(e.what()));
            MessageBox(hwnd, errorMessage.c_str(), L"Error", MB_ICONERROR);
            return 0;
        }

        if (output.empty()) {
            MessageBox(hwnd, L"Error: No output from the command.", L"Error", MB_ICONERROR);
            return 0;
        }

        if (output.at(0) == '0') {
            SetWindowText(GetDlgItem(hwnd, 2), L"Unset Confidential");
            SetWindowText(GetDlgItem(hwnd, 3), L"This document is confidential.");

            command = exePath + L" toggle";
        }
        else if (output.at(0) == '1') {
            SetWindowText(GetDlgItem(hwnd, 2), L"Set Confidential");
            SetWindowText(GetDlgItem(hwnd, 3), L"");

            command = exePath + L" toggle";
        }
        else {
            std::wstring errorMessage = L"Unexpected output: " + std::wstring(output.begin(), output.end());
            MessageBox(hwnd, errorMessage.c_str(), L"Error", MB_ICONERROR);
            return 0;
        }

        try {
            RunCommand(command);
        }
        catch (const std::exception& e) {
            std::wstring errorMessage = L"Error running command: " + std::wstring(e.what(), e.what() + strlen(e.what()));
            MessageBox(hwnd, errorMessage.c_str(), L"Error", MB_ICONERROR);
            return 0;
        }

        return 0;
    }
    break;
  case WM_CREATE: {
    SetTimer(hwnd, 10, 100, NULL);
    HWND hButton = CreateWindowW(L"BUTTON", L"", WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON | BS_ICON | BS_OWNERDRAW,
      160, 10, 30, 30, hwnd, (HMENU)1, hInst, NULL);
    // Set the icon on the button
    // SendMessage(hButton, BM_SETIMAGE, IMAGE_ICON, (LPARAM)hIcon);
    CreateWindowW(L"BUTTON", L"", WS_TABSTOP | WS_CHILD | BS_DEFPUSHBUTTON | BS_OWNERDRAW | WS_VISIBLE, 20, 60, 160, 30, hwnd, (HMENU)2, hInst, NULL);
    CreateWindow(L"static", L"", WS_CHILD | WS_TABSTOP | SS_OWNERDRAW | WS_VISIBLE, 20, 110, 160, 40, hwnd, (HMENU)3, hInst, NULL);
    // Create a brush with the desired background color
    
    break;
  }
  case WM_TIMER: {
    if (!IsWindow(targetWnd)) {
      KillTimer(hwnd, 100);
      DestroyWindow(hwnd);
    }
    TCHAR btnText[256];
    GetWindowText(GetDlgItem(hwnd, 2), btnText, 256);
    if (_tcscmp(btnText, L"") == 0) {
        std::wstring exeDirectory = GetExecutableDirectory();
        std::wstring exeName = L"WordDocumentUpdater.exe";
        std::wstring exePath = exeDirectory + L"\\" + exeName;
        exePath = L"\"" + exePath + L"\"";

        std::wstring command = exePath + L" get";
        std::string output;

        try {
            output = RunCommand(command);
        }
        catch (const std::exception& e) {
            std::wstring errorMessage = L"Error running command: " + std::wstring(e.what(), e.what() + strlen(e.what()));
            MessageBox(hwnd, errorMessage.c_str(), L"Error", MB_ICONERROR);
            return 0;
        }

        if (output.empty()) {
            SetWindowText(GetDlgItem(hwnd, 2), L"Set Confidential");
            SetWindowText(GetDlgItem(hwnd, 3), L"");
        }
        else if (output.at(0) == '1') {
            SetWindowText(GetDlgItem(hwnd, 2), L"Unset Confidential");
            SetWindowText(GetDlgItem(hwnd, 3), L"This document is confidential.");
        }
        else if (output.at(0) == '0') {
            SetWindowText(GetDlgItem(hwnd, 2), L"Set Confidential");
            SetWindowText(GetDlgItem(hwnd, 3), L"");
        }
    }
    return 0;
  }
  break;
  case WM_PAINT: {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hwnd, &ps);
    // Create a compatible device context for double buffering
    HDC hdcMem = CreateCompatibleDC(hdc);
    HBITMAP hBitmap = CreateCompatibleBitmap(hdc, 200, rect.bottom - rect.top);
    SelectObject(hdcMem, hBitmap);
    HBRUSH whiteBrush = CreateSolidBrush(RGB(246, 247, 248));
    RECT backRect = { 0, 0, 200, rect.bottom - rect.top };
    FillRect(hdcMem, &backRect, whiteBrush);
    // Copy contents of the off-screen bitmap to the window's device context
    BitBlt(hdc, 0, 0, 200, rect.bottom - rect.top, hdcMem, 0, 0, SRCCOPY);
    // Cleanup
    DeleteDC(hdcMem);
    DeleteObject(hBitmap);
    EndPaint(hwnd, &ps);
    break;
  }
  break;
  case WM_DRAWITEM: {
    DRAWITEMSTRUCT* pdis = (DRAWITEMSTRUCT*)lParam;
    if (pdis->CtlID == 1) {
      SetBkMode(pdis->hDC, TRANSPARENT);
      HICON hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_CLOSE));
      DrawIcon(pdis->hDC, 0, 0, hIcon);
    }
    if (pdis->CtlID == 2) {
      HDC hdcMem = CreateCompatibleDC(pdis->hDC);
      HBITMAP hBitmap = CreateCompatibleBitmap(pdis->hDC, 200, rect.bottom - rect.top);
      SelectObject(hdcMem, hBitmap);
      RECT rc = pdis->rcItem;
      SetBkMode(hdcMem, TRANSPARENT);
      // Fill the background
      HBRUSH brBackground = CreateSolidBrush(RGB(56, 65, 74));
      FillRect(hdcMem, &rc, brBackground);
      DeleteObject(brBackground);
      // Draw the text
      SetTextColor(hdcMem, RGB(255, 255, 255));
      TCHAR btnText[256];
      GetWindowText(GetDlgItem(hwnd, 2), btnText, 256);
      DrawText(hdcMem, btnText, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
      // Copy contents of the off-screen bitmap to the window's device context
      BitBlt(pdis->hDC, 0, 0, 160, 30, hdcMem, 0, 0, SRCCOPY);
      // Cleanup
      DeleteDC(hdcMem);
      DeleteObject(hBitmap);

      if (pdis->itemState & ODS_SELECTED) {
        // Add effect when the button is pressed
      }
    }
    if (pdis->CtlID == 3) {
      HDC hdcMem = CreateCompatibleDC(pdis->hDC);
      HBITMAP hBitmap = CreateCompatibleBitmap(pdis->hDC, 200, rect.bottom - rect.top);
      SelectObject(hdcMem, hBitmap);
      RECT rc = pdis->rcItem;
      SetBkMode(hdcMem, TRANSPARENT);
      // Fill the background
      HBRUSH brBackground = CreateSolidBrush(RGB(246, 247, 248));
      FillRect(hdcMem, &rc, brBackground);
      DeleteObject(brBackground);

      // Draw the text
      SetTextColor(hdcMem, RGB(255, 0, 0));
      TCHAR btnText[256];
      GetWindowText(GetDlgItem(hwnd, 3), btnText, 256);
      DrawText(hdcMem, btnText, -1, &rc, DT_CENTER | DT_VCENTER | DT_WORDBREAK);
      // Copy contents of the off-screen bitmap to the window's device context
      BitBlt(pdis->hDC, 0, 0, 160, 30, hdcMem, 0, 0, SRCCOPY);
      // Cleanup
      DeleteDC(hdcMem);
      DeleteObject(hBitmap);
      
    }
  }
  break;
  case WM_CLOSE:
    KillTimer(hwnd, 100);
    DestroyWindow(hwnd);
    break;
  default:
    return DefWindowProc(hwnd, msg, wParam, lParam);
  }
  return 0;
}

// Function to create the tray icon
void CreateTrayIcon(HWND hwnd) {
  memset(&g_notifyIconData, 0, sizeof(NOTIFYICONDATA));
  g_notifyIconData.cbSize = sizeof(NOTIFYICONDATA);
  g_notifyIconData.hWnd = hwnd;
  g_notifyIconData.uID = 1;
  g_notifyIconData.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
  g_notifyIconData.uCallbackMessage = WM_TRAYICON;
  g_notifyIconData.hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_CONFIDENTIAL));
  _tcscpy_s(g_notifyIconData.szTip, _T("Tray Icon App"));

  Shell_NotifyIcon(NIM_ADD, &g_notifyIconData);
}

// Function to remove the tray icon
void RemoveTrayIcon() {
  Shell_NotifyIcon(NIM_DELETE, &g_notifyIconData);
}

// Function to show the context menu
void ShowContextMenu(HWND hwnd, POINT pt) {
  HMENU hMenu = CreatePopupMenu();
  if (hMenu) {
    InsertMenu(hMenu, -1, MF_BYPOSITION, ID_TRAY_EXIT, _T("Exit"));

    SetForegroundWindow(hwnd);  // Must call this before TrackPopupMenu
    TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hwnd, NULL);
    DestroyMenu(hMenu);
  }
}

// Function to update the button position
void UpdateButtonPosition() {
  EnumWindows(EnumWindowsProc, NULL);
}

// EnumWindows callback function
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam) {
  DWORD processID;
  GetWindowThreadProcessId(hwnd, &processID);
  std::wstring processName = GetProcessName(processID);

  TCHAR windowClass[256];
  GetClassName(hwnd, windowClass, 256);

  if (IsTargetProcess(processName)) {
    if (_tcscmp(windowClass, _T("OpusApp")) == 0 ||
        _tcscmp(windowClass, _T("PPTFrameClass")) == 0 ||
        _tcscmp(windowClass, _T("Framework::CFrame")) == 0 ||
        _tcscmp(windowClass, _T("XLMAIN")) == 0 ||
        _tcscmp(windowClass, _T("AcrobatSDIWindow")) == 0) {
        CreateOrUpdateButton(hwnd);
      }
  }

  return TRUE;
}

// Function to check if the process is one of the target processes
bool IsTargetProcess(const std::wstring& processName) {
  return g_targetProcesses.find(processName) != g_targetProcesses.end();
}

// Function to get the process name from the process ID
std::wstring GetProcessName(DWORD processID) {
  std::wstring processName;
  HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processID);
  if (hProcess) {
    TCHAR buffer[MAX_PATH];
    if (GetModuleBaseName(hProcess, NULL, buffer, MAX_PATH)) {
      processName = buffer;
    }
    CloseHandle(hProcess);
  }
  return processName;
}

// Function to create or update the button for a target window
void CreateOrUpdateButton(HWND targetWnd) {
  RECT rect;
  GetWindowRect(targetWnd, &rect);
  int buttonWidth = 40;
  int buttonHeight = 32;
  int buttonX = rect.left;
  int buttonY = (rect.top + rect.bottom) / 2 - buttonHeight / 2;

  if (g_buttonMap.find(targetWnd) == g_buttonMap.end()) {
    
    HWND buttonWnd = CreateWindowEx(
      WS_EX_TOPMOST | WS_EX_TOOLWINDOW,          // Extended window style to hide from taskbar
      _T("ButtonWindowClass"),   // Window class name
      _T(""),                    // Window title
      WS_POPUP | WS_VISIBLE,     // Window style
      buttonX, buttonY, buttonWidth, buttonHeight,
      NULL, NULL, hInst, NULL
    );
    ShowWindow(buttonWnd, SW_HIDE);

    HWND paneWnd = CreateWindowEx(
      WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
      _T("PaneWindowClass"),
      _T(""),
      WS_POPUP,
      rect.left, rect.top, 200, rect.bottom - rect.top, NULL, NULL, hInst, NULL
    );
    ShowWindow(paneWnd, SW_HIDE);

    g_buttonMap[targetWnd] = buttonWnd;
    g_panMap[buttonWnd] = paneWnd;
    g_targetMap[paneWnd] = targetWnd;

    DWORD processID;
    GetWindowThreadProcessId(targetWnd, &processID);
    int process_id = static_cast<int>(processID);
    std::wstring commandLine = GetCommandLineFromProcess(process_id);
    g_pathMap[paneWnd] = GetSubstringFromLastQuote(commandLine);
  }
  else {
    HWND buttonWnd = g_buttonMap[targetWnd];
    HWND paneWnd = g_panMap[buttonWnd];
    HWND currentWnd = GetForegroundWindow();
    if(currentWnd != buttonWnd && currentWnd != targetWnd && currentWnd != paneWnd) {
      ShowWindow(buttonWnd, SW_HIDE);
      ShowWindow(paneWnd, SW_HIDE);
    }
    else {
      if(!g_checkMap[buttonWnd]) SetWindowPos(buttonWnd, HWND_TOPMOST, buttonX, buttonY, buttonWidth, buttonHeight, SWP_NOZORDER | SWP_SHOWWINDOW);
      else SetWindowPos(paneWnd, HWND_TOPMOST, rect.left, rect.top, 200, rect.bottom - rect.top, SWP_NOZORDER | SWP_SHOWWINDOW);
    }
  }
}