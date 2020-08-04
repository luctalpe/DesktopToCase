// Win32 Dialog.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "resource.h"

#include <iostream>

#pragma once

#include <windows.h>
#include <Windowsx.h>
#include <winstring.h>
#include <strsafe.h>
#include <shobjidl.h>
#include <uxtheme.h>
#include <wrl/implements.h>
#include <stdlib.h>
#include <malloc.h>
#include <memory.h>
#include <tchar.h>
#include <regex>
#include <map>
#include <vector>
#include <psapi.h>
#include <winternl.h>
#include <processsnapshot.h>
#include <objbase.h>
#include <ObjectArray.h>
#include <Hstring.h>


#define SWM_EXIT	WM_APP + 1
#define SWM_STARTAUTOMATICALLY WM_APP + 2
#define SWM_SHOWDEBUG WM_APP + 3 //	show the window
#define SWM_ASSIGNUPDATEALLCASESTODESK  SWM_SHOWDEBUG+3
#define SWM_REMOVEALLDESKTOPS   SWM_SHOWDEBUG+4
#define SWM_MOVETOCLIPBOARDESKTOP      SWM_SHOWDEBUG+5
#define SWM_MOVETODESKTOP SWM_SHOWDEBUG+6

WCHAR MyAppKey[] = L"DesktopForCase";
WCHAR ModuleFileName[MAX_PATH+1];
bool caseInSensStringCompareCpp11(std::wstring& str1, std::wstring& str2)
{
    return ((str1.size() == str2.size()) && std::equal(str1.begin(), str1.end(), str2.begin(), [](wchar_t& c1, wchar_t& c2) {
        return (c1 == c2 || std::toupper(c1) == std::toupper(c2));
        }));
}

 
BOOL IsRunAtStartup( BOOL bRemove = FALSE , BOOL bSet = FALSE ) {
    HKEY hKey ;
    LONG lnRes = RegOpenKeyExW(        HKEY_CURRENT_USER ,        L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run",        0L, KEY_WRITE| KEY_QUERY_VALUE | KEY_SET_VALUE ,       &hKey            // address of handle of open key 
    );
    BOOL bReturn = FALSE;
    WCHAR KeyValue[MAX_PATH+1];
    GetModuleFileNameW(NULL, ModuleFileName, MAX_PATH);
    if (lnRes == ERROR_SUCCESS) {
        DWORD dwType = 0;
        DWORD cbData = sizeof(KeyValue);
        WCHAR *Buffer = NULL;
        LONG Status;
        Status = RegQueryValueEx(hKey, MyAppKey, NULL, &dwType, (LPBYTE)KeyValue, &cbData);
        if (Status == ERROR_SUCCESS && dwType == REG_SZ) {
            std::wstring regvalue((WCHAR*)KeyValue);
            std::wstring mymodule(ModuleFileName);
            if (caseInSensStringCompareCpp11(regvalue, mymodule)) {
                bReturn = TRUE;
                if (bRemove) {
                    RegDeleteValueW(hKey, MyAppKey);
                }
            }         
        }
        if (bSet) {
            RegSetKeyValueW(hKey, NULL, MyAppKey, REG_SZ, ModuleFileName, (DWORD)wcslen(ModuleFileName) * 2 + 2);
        }
        RegCloseKey(hKey);
    }
    
    return bReturn;
}

using namespace Microsoft::WRL;
using namespace std;

#pragma comment (lib, "onecore.lib" )

#pragma region Virtual_Desktop_interface

const CLSID CLSID_ImmersiveShell = {
    0xC2F03A33, 0x21F5, 0x47FA, 0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39 };

const CLSID CLSID_VirtualDesktopAPI_Unknown = {
    0xC5E0CDCA, 0x7B6E, 0x41B2, 0x9F, 0xC4, 0xD9, 0x39, 0x75, 0xCC, 0x46, 0x7B };

const IID IID_IVirtualDesktopManagerInternal = {
    0xEF9F1A6C, 0xD3CC, 0x4358, 0xB7, 0x12, 0xF8, 0x4B, 0x63, 0x5B, 0xEB, 0xE7 };

const CLSID CLSID_IVirtualNotificationService = {
    0xA501FDEC, 0x4A09, 0x464C, 0xAE, 0x4E, 0x1B, 0x9C, 0x21, 0xB8, 0x49, 0x18 };

EXTERN_C const IID  IID_IWin32ApplicationView;
MIDL_INTERFACE("F3527610-C76D-4316-AC8E-28651ACF3DF3")
IWin32ApplicationView : IUnknown
{
    virtual HRESULT GetWindow(HWND * pHwnd) = 0;
    virtual HRESULT GetCloakableWindows(UINT* pWindowCount, HWND** ppWindows) = 0;
};


EXTERN_C const IID  IID_IApplicationView;
MIDL_INTERFACE("372E1D3B-38D3-42E4-A15B-8AB2B178F513")
IApplicationView : IInspectable
{
   virtual HRESULT SetFocus() = 0;
    virtual HRESULT SwitchTo() = 0;
    virtual HRESULT TryInvokeBack(IUnknown* pCallback) = 0;
    virtual HRESULT GetThumbnailWindow(HWND* pThumbnailWindow) = 0;
    virtual HRESULT GetMonitor(IUnknown** ppMonitor) = 0;
    virtual HRESULT GetVisibility(BOOL* pfIsVisible) = 0;
    virtual HRESULT SetCloak(int  type, BOOL fCloak) = 0;
    virtual HRESULT GetPosition(REFIID riid, void** ppv) = 0; // We always support IApplicationViewPosition
    virtual HRESULT SetPosition(IUnknown* pPosition) = 0;
    virtual HRESULT InsertAfterWindow(HWND hwndInsertAfter) = 0;
    virtual HRESULT GetExtendedFramePosition(RECT* prcExtendedFrame) = 0; // Get the 1px border-sized window rect.
    virtual HRESULT GetAppUserModelId(LPWSTR* ppszAppUserModelId) = 0; // *ppszAppUserModelId must be freed with CoTaskMemFree.
    virtual HRESULT SetAppUserModelId(LPCWSTR pszAppUserModelId) = 0;
    virtual   HRESULT IsEqualByAppUserModelId(LPCWSTR pszAppUserModelId,  BOOL* pfEqual) = 0;
    virtual HRESULT GetViewState(UINT* pnShowCmd) = 0;
    virtual HRESULT SetViewState(UINT nShowCmd) = 0;
    virtual HRESULT GetNeediness(BOOL* pfIsNeedy) = 0;
    virtual HRESULT GetLastActivationTimestamp(ULONGLONG* pullTime) = 0;
    virtual  HRESULT SetLastActivationTimestamp(ULONGLONG ullTime) = 0;
    virtual HRESULT GetVirtualDesktopId(LPGUID pGuid) = 0;
    virtual  HRESULT SetVirtualDesktopId(REFGUID refGuid) = 0;
    virtual HRESULT GetShowInSwitchers(BOOL* pfShow) = 0;
    virtual HRESULT SetShowInSwitchers(BOOL fShow) = 0;
    virtual HRESULT GetScaleFactor(UINT* pScaleFactor) = 0;
    virtual HRESULT CanReceiveInput(BOOL* pfEnabled) = 0;
    virtual HRESULT GetCompatibilityPolicyType(IUnknown* pCompatPolicy) = 0;
    virtual HRESULT SetCompatibilityPolicyType(IUnknown compatPolicy) = 0;
    virtual HRESULT GetPositionerPriority(IUnknown** ppPriority) = 0;
    virtual HRESULT SetPositionerPriority(IUnknown* pPriority) = 0;
    virtual HRESULT GetSizeConstraints(IUnknown* pMonitor, SIZE* pMinSize, SIZE* pMaxSize) = 0;
    virtual HRESULT GetSizeConstraintsForDpi(UINT dpi, SIZE* pMinSize, SIZE* pMaxSize) = 0;
    virtual HRESULT SetSizeConstraintsForDpi(UINT const* pDpi,  SIZE const* pMinSize, SIZE const* pMaxSize) = 0;
    virtual HRESULT QuerySizeConstraintsFromApp() = 0;
    virtual HRESULT OnMinSizePreferencesUpdated(HWND hwndApp) = 0;
    virtual HRESULT ApplyOperation(IUnknown* pOperation) = 0;
    virtual HRESULT IsTray(BOOL* isTray) = 0;
    virtual HRESULT IsInHighZOrderBand(BOOL* isInHighZOrderBand) = 0;
    virtual HRESULT IsSplashScreenPresented(BOOL* isSplashScreenPresented) = 0;
    virtual HRESULT Flash() = 0;
    virtual HRESULT GetRootSwitchableOwner(IApplicationView** rootOwner) = 0;
    virtual HRESULT EnumerateOwnershipTree(IObjectArray** ownershipChain) = 0;
    virtual HRESULT GetEnterpriseId(LPWSTR* ppszEnterpriseId) = 0; // *ppszEnterpriseId must be freed with CoTaskMemFree.
    virtual HRESULT GetEnterpriseChromePreference(BOOL* shouldShowEnterpriseChrome) = 0;
    virtual HRESULT IsMirrored(BOOL* pfIsRTL) = 0;
    virtual HRESULT GetFrameworkViewType(IUnknown* pViewType) = 0;
    virtual HRESULT GetCanTab(BOOL* canTab) = 0;
    virtual HRESULT SetCanTab(BOOL canTab) = 0;
    virtual HRESULT GetIsTabbed(BOOL* isTabbed) = 0;
    virtual HRESULT SetIsTabbed(BOOL isTabbed) = 0;
    virtual HRESULT RefreshCanTab() = 0;
    virtual HRESULT GetIsOccluded(BOOL* isOccluded) = 0;
    virtual HRESULT SetIsOccluded(BOOL isOccluded) = 0;
    virtual HRESULT UpdateEngagementFlags(int flagsToUpdate, int  flagValues) = 0;
    virtual HRESULT SetForceActiveWindowAppearance(BOOL value) = 0;
    virtual HRESULT GetLastActivationFILETIME(FILETIME* value) = 0;
    virtual HRESULT GetPersistingStateName(LPWSTR* persistingStateName) = 0;
};


EXTERN_C const IID  IID_IApplicationViewCollection;
MIDL_INTERFACE("1841C6D7-4F9D-42C0-AF41-8747538F10E5")
IApplicationViewCollection : IUnknown
{

    virtual HRESULT GetViews(IObjectArray * *ppViews) = 0;
    virtual HRESULT GetViewsByZOrder(IObjectArray** ppViews) = 0;
    virtual HRESULT GetViewsByAppUserModelId(LPCWSTR pszAppUserModelId,  IObjectArray** ppViews) = 0;

    virtual HRESULT GetViewForHwnd(HWND hwnd,  IApplicationView** ppView) = 0;
    // virtual HRESULT GetViewForApplication([in] IImmersiveApplication* pApp, [out] IApplicationView** ppView) = 0;  
    virtual HRESULT GetViewForApplication(IUnknown* pApp,  IApplicationView** ppView) = 0;
    virtual HRESULT GetViewForAppUserModelId(LPCWSTR pszAppUserModelId, IApplicationView** ppView) = 0;
    virtual HRESULT GetViewInFocus(IApplicationView** ppView) = 0;
    virtual  HRESULT TryGetLastActiveVisibleView(IApplicationView** value) = 0;
    virtual HRESULT RefreshCollection() = 0;
    // virtual HRESULT RegisterForApplicationViewChanges( IApplicationViewChangeListener* pListener, [out] DWORD* pdwCookie) = 0;
    virtual HRESULT RegisterForApplicationViewChanges(IUnknown* pListener,  DWORD* pdwCookie) = 0;
    virtual HRESULT UnregisterForApplicationViewChanges(DWORD dwCookie) = 0;
};



EXTERN_C const IID IID_IVirtualDesktop;

MIDL_INTERFACE("FF72FFDD-BE7E-43FC-9C03-AD81681E88E4")
IVirtualDesktop : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE IsViewVisible(
        IApplicationView * pView,
        int* pfVisible) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetID(
        GUID* pGuid) = 0;
};

enum AdjacentDesktop
{
    LeftDirection = 3,
    RightDirection = 4
};

EXTERN_C const IID IID_IVirtualDesktopManagerInternal;

MIDL_INTERFACE("f31574d6-b682-4cdc-bd56-1827860abec6")
IVirtualDesktopManagerInternal : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE GetCount(
        UINT * pCount) = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveViewToDesktop(
        IApplicationView* pView,
        IVirtualDesktop* pDesktop) = 0;

    // 10240
    virtual HRESULT STDMETHODCALLTYPE CanViewMoveDesktops(
        IApplicationView* pView,
        int* pfCanViewMoveDesktops) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetCurrentDesktop(
        IVirtualDesktop** desktop) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetDesktops(
        IObjectArray** ppDesktops) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetAdjacentDesktop(
        IVirtualDesktop* pDesktopReference,
        AdjacentDesktop uDirection,
        IVirtualDesktop** ppAdjacentDesktop) = 0;

    virtual HRESULT STDMETHODCALLTYPE SwitchDesktop(
        IVirtualDesktop* pDesktop) = 0;

    virtual HRESULT STDMETHODCALLTYPE CreateDesktopW(
        IVirtualDesktop** ppNewDesktop) = 0;

    virtual HRESULT STDMETHODCALLTYPE RemoveDesktop(
        IVirtualDesktop* pRemove,
        IVirtualDesktop* pFallbackDesktop) = 0;

    // 10240
    virtual HRESULT STDMETHODCALLTYPE FindDesktop(
        GUID* desktopId,
        IVirtualDesktop** ppDesktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDesktopSwitchIncludeExcludeViews(
        /* [in] */ __RPC__in_opt IVirtualDesktop* pTargetDesktop,
        /* [out] */ __RPC__deref_out_opt IObjectArray** ppIncludeViews,
        /* [out] */ __RPC__deref_out_opt IObjectArray** ppExcludeViews) = 0;

};

EXTERN_C const IID IID_IVirtualDesktopManagerInternal2;


MIDL_INTERFACE("0f3a72b0-4566-487e-9a33-4ed302f6d6ce")
IVirtualDesktopManagerInternal2 : public IVirtualDesktopManagerInternal
{
public:
    virtual HRESULT STDMETHODCALLTYPE SetDesktopName(IVirtualDesktop * desktop,  HSTRING name) = 0;

};


EXTERN_C const IID IID_IVirtualDesktopNotification;

MIDL_INTERFACE("C179334C-4295-40D3-BEA1-C654D965605A")
IVirtualDesktopNotification : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE VirtualDesktopCreated(
        IVirtualDesktop * pDesktop) = 0;

    virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyBegin(
        IVirtualDesktop* pDesktopDestroyed,
        IVirtualDesktop* pDesktopFallback) = 0;

    virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyFailed(
        IVirtualDesktop* pDesktopDestroyed,
        IVirtualDesktop* pDesktopFallback) = 0;

    virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyed(
        IVirtualDesktop* pDesktopDestroyed,
        IVirtualDesktop* pDesktopFallback) = 0;

    virtual HRESULT STDMETHODCALLTYPE ViewVirtualDesktopChanged(
        IApplicationView* pView) = 0;

    virtual HRESULT STDMETHODCALLTYPE CurrentVirtualDesktopChanged(
        IVirtualDesktop* pDesktopOld,
        IVirtualDesktop* pDesktopNew) = 0;

};

EXTERN_C const IID IID_IVirtualDesktopNotificationService;

MIDL_INTERFACE("0CD45E71-D927-4F15-8B0A-8FEF525337BF")
IVirtualDesktopNotificationService : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE Register(
        IVirtualDesktopNotification * pNotification,
        DWORD * pdwCookie) = 0;

    virtual HRESULT STDMETHODCALLTYPE Unregister(
        DWORD dwCookie) = 0;
};

#pragma endregion Virtual_Desktop_interface

// Utility function to translate a GUID to wstring
std::wstring  GetGuid(const GUID& guid)
{
    std::wstring guidStr(40, L'\0');
    ::StringFromGUID2(guid, const_cast<LPOLESTR>(guidStr.c_str()), (DWORD)guidStr.length());
    return guidStr;
}

#pragma region Sample_Virtual_API_functions 
HRESULT EnumVirtualDesktops(IVirtualDesktopManagerInternal* pDesktopManager)
{
    //  std::wcout << L"<<< EnumDesktops >>>" << std::endl;

    IObjectArray* pObjectArray = nullptr;
    HRESULT hr = pDesktopManager->GetDesktops(&pObjectArray);

    if (SUCCEEDED(hr))
    {
        UINT count;
        hr = pObjectArray->GetCount(&count);

        if (SUCCEEDED(hr))
        {
            std::wcout << L"Count: " << count << std::endl;

            for (UINT i = 0; i < count; i++)
            {
                IVirtualDesktop* pDesktop = nullptr;

                if (FAILED(pObjectArray->GetAt(i, __uuidof(IVirtualDesktop), (void**)&pDesktop)))
                    continue;

                GUID id = { 0 };
                if (SUCCEEDED(pDesktop->GetID(&id)))
                {
                    std::wcout << L"\t #" << i << L": ";
                    std::wcout << GetGuid(id);
                    std::wcout << std::endl;
                }

                pDesktop->Release();
            }
        }

        pObjectArray->Release();
    }

    std::wcout << std::endl;
    return hr;
}

HRESULT EnumAdjacentDesktops(IVirtualDesktopManagerInternal* pDesktopManager)
{
    std::wcout << L"<<< EnumAdjacentDesktops >>>" << std::endl;

    IVirtualDesktop* pDesktop = nullptr;
    HRESULT hr = pDesktopManager->GetCurrentDesktop(&pDesktop);

    if (SUCCEEDED(hr))
    {
        GUID id = { 0 };
        IVirtualDesktop* pAdjacentDesktop = nullptr;
        hr = pDesktopManager->GetAdjacentDesktop(pDesktop, AdjacentDesktop::LeftDirection, &pAdjacentDesktop);

        std::wcout << L"At left direction: ";

        if (SUCCEEDED(hr))
        {
            if (SUCCEEDED(pAdjacentDesktop->GetID(&id)))
                // PrintGuid(id);

                pAdjacentDesktop->Release();
        }
        else
            std::wcout << L"NULL";
        std::wcout << std::endl;

        id = { 0 };
        pAdjacentDesktop = nullptr;
        hr = pDesktopManager->GetAdjacentDesktop(pDesktop, AdjacentDesktop::RightDirection, &pAdjacentDesktop);

        std::wcout << L"At right direction: ";

        if (SUCCEEDED(hr))
        {
            if (SUCCEEDED(pAdjacentDesktop->GetID(&id)))
                // PrintGuid(id);

                pAdjacentDesktop->Release();
        }
        else
            std::wcout << L"NULL";
        std::wcout << std::endl;

        pDesktop->Release();
    }

    std::wcout << std::endl;
    return hr;
}

HRESULT ManageVirtualDesktops(IVirtualDesktopManagerInternal* pDesktopManager)
{
    std::wcout << L"<<< ManageVirtualDesktops >>>" << std::endl;
    std::wcout << L"Sleep period: 2000 ms" << std::endl;

    ::Sleep(2000);


    IVirtualDesktop* pDesktop = nullptr;
    HRESULT hr = pDesktopManager->GetCurrentDesktop(&pDesktop);

    if (FAILED(hr))
    {
        std::wcout << L"\tFAILED can't get current desktop" << std::endl;
        return hr;
    }

    std::wcout << L"Creating desktop..." << std::endl;

    IVirtualDesktop* pNewDesktop = nullptr;
    hr = pDesktopManager->CreateDesktopW(&pNewDesktop);

    if (SUCCEEDED(hr))
    {
GUID id;
hr = pNewDesktop->GetID(&id);

if (FAILED(hr))
{
    std::wcout << L"\tFAILED GetID" << std::endl;
    pNewDesktop->Release();
    return hr;
}

std::wcout << L"\t";
// PrintGuid(id);
std::wcout << std::endl;

std::wcout << L"Switching to desktop..." << std::endl;
hr = pDesktopManager->SwitchDesktop(pNewDesktop);

if (FAILED(hr))
{
    std::wcout << L"\tFAILED SwitchDesktop" << std::endl;
    pNewDesktop->Release();
    return hr;
}

::Sleep(2000);

std::wcout << L"Removing desktop..." << std::endl;

if (SUCCEEDED(hr))
{
    hr = pDesktopManager->RemoveDesktop(pNewDesktop, pDesktop);
    pDesktop->Release();

    if (FAILED(hr))
    {
        std::wcout << L"\tFAILED RemoveDesktop" << std::endl;
        pNewDesktop->Release();
        return hr;
    }
}
    }

    std::wcout << std::endl;
    return hr;
}

#pragma endregion  Sample_Virtual_API_functions

wregex CaseId(L"[1][12][0-9]{13}");

// Will manage Virtual Desktop /CaseId association 
class VirtualDesktopCase {
    std::map<std::wstring, GUID> m_pList;                           // Dictionnary containing the relation between virtal desktop (GUID) and caseid 

    std::vector< std::pair<GUID, DWORD>> m_pDesktopList;           // Dictionnary for each Virtual Desktop will keep the deskop number and if in use (any top window ?) 

    std::wstring GetCaseId(GUID& g) {
        std::map<std::wstring, GUID>::iterator it = m_pList.begin();
        while (it != m_pList.end()) {
            if (IsEqualGUID(it->second, g)) {
                return it->first;
            }
            it++;
        }
        return L"";
    }
public:
  
    void Reset() {
        m_pList.clear();
        m_pDesktopList.clear();
    }
    void PrintDesktops(HMENU h = NULL, UINT* pPos = NULL) {
        for (int i = 0; i < m_pDesktopList.size(); i++) {
            std::pair<GUID, DWORD>& d = m_pDesktopList[i];
            wcout << L"Desktop " << (i + 1) << L" " << GetGuid(d.first) << L" : " << GetCaseId(d.first) << endl;
            if (h != NULL && pPos != NULL) {
                MENUITEMINFOW mi;
                mi.cbSize = sizeof(MENUITEMINFO);
                mi.fMask = MIIM_TYPE | MIIM_ID;
                mi.fType = MFT_STRING;

                mi.wID = (UINT)SWM_MOVETODESKTOP+i;
                WCHAR iw[10];
                // _itow(i + 1, iw, 10);
                _itow_s(i + 1, iw, 10);
                wstring s(L"Move to Desktop");
                s += wstring(iw) + wstring(L" ") + wstring(GetCaseId(d.first));

                wchar_t* ptr = _wcsdup(s.c_str());
                mi.dwTypeData = ptr;

                InsertMenuItemW(h, (*pPos)++, TRUE, &mi);
                free(ptr);
            }
        }
    }


    BOOL AddDesktop(GUID Desktop) {
        for (int i = 0; i < m_pDesktopList.size(); i++) {
            std::pair<GUID, DWORD>& d = m_pDesktopList[i];
            if (IsEqualGUID(d.first, Desktop))  {
            return true;
            }
        }
        m_pDesktopList.push_back(std::pair< GUID, DWORD>(Desktop, 0));
        return true;
    }
    void SetDesktopInUse(GUID g) {
        for (int i = 1; i < m_pDesktopList.size(); i++) {
            std::pair<GUID, DWORD> var = m_pDesktopList[i];
            if (IsEqualGUID(var.first, g)) {
                if (var.second == 0) {
                    var.second = 1;
                    m_pDesktopList[i] = var;
                    break;
                }
            }
        }
    }
    // 4 helper methods to manager the vDesktop/cases map
    GUID GetDesktopId(std::wstring CaseId) {                            // Find the vdesktop associated to a case in m_pList dictionary : renturn null GUID if not
        auto result = m_pList.find(CaseId);
        if (result == m_pList.end()) return { 0 };
        return result->second;      // Return Desktop GUID
    }
    GUID GetDesktopGuid(int i) {
        std::pair<GUID, DWORD>& d = m_pDesktopList[i];
        return d.first;
    }
    void AddDesktopId(wstring CaseId, GUID DesktopId) {
        // wcout << L"VirtualDesktopCase::AddDesktopId( " << CaseId << L"," << GetGuid(DesktopId) << L")"<< endl;
        m_pList[CaseId] = DesktopId;
    }
    /* void RemoveCaseId(wstring CaseId) {
        auto result = m_pList.find(CaseId);
        if (result != m_pList.end()) {
            m_pList.erase(result);
        }
    }*/
    BOOL IsDesktopIdAlreadyInUse(GUID DeskstopId) {     // Will check if we already reference this deskop in our desktop/casId list 
        std::map<std::wstring, GUID>::iterator it = m_pList.begin();

        while (it != m_pList.end()) {
            if (IsEqualGUID(it->second, DeskstopId))
                return true;
            it++;
        }
        return false;
    }

};

std::wstring GetClipboardText()
{
    WCHAR* pszText = L"";
    // Try opening the clipboard
    if (!OpenClipboard(nullptr))
        return pszText;

    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData == nullptr) {
        CloseClipboard();
        return pszText;
    }
    // Lock the handle to get the actual text pointer
    pszText = static_cast<WCHAR*>(GlobalLock(hData));
  
    // Release the lock
    GlobalUnlock(hData);

    // Release the clipboard
    CloseClipboard();

    return pszText;
}

class LTDesktopManager {
    VirtualDesktopCase DesktopCase;     // Will keep desktopId map created to manage  cases ...
    GUID m_CurrentDesktopGuid = { 0 };
    IServiceProvider* m_pServiceProvider;
    IVirtualDesktopManagerInternal* m_pDesktopManagerInternal;
    IVirtualDesktopManager* m_pDesktopManager;
    IVirtualDesktopManagerInternal2* m_pDesktopManagerInternal2;
    IApplicationViewCollection* m_pApplicationViewCollection;

    map<DWORD, DWORD> m_ProcessAlreadyProcessed;        // Will remember during enumprocess if a Process has already a top window matching a case = 1 or a file matching a case = 2

    HRESULT MoveToDesktop(HWND hwnd, GUID HwndDesktopId, wstring CaseId) {         // Will move a top window to the virtual desktop hosting "CaseId"
        IVirtualDesktop* pVirtualDesktop = nullptr;
        HRESULT hr;
        GUID g = DesktopCase.GetDesktopId(CaseId);      // Do we already have a virtual desktop assigned for this CaseId

        if (IsEqualGUID(g, { 0 })) {                    // No vDesktop for this caseid yet
            hr = m_pDesktopManagerInternal->CreateDesktopW(&pVirtualDesktop); // Will create the desktop - TO DO: check if empty desktop exist
            if (FAILED(hr)) return hr;

            hr = pVirtualDesktop->GetID(&g);        // get the GUID for this new virtual desktop 
            DesktopCase.AddDesktop(g);
            DesktopCase.AddDesktopId(CaseId, g);   // remember the DesktopId for this caseid  in our map 
            DesktopCase.SetDesktopInUse(g);


            IApplicationView* pApplicationView = nullptr;
            hr = this->m_pApplicationViewCollection->GetViewForHwnd(hwnd, &pApplicationView);  // will get the desktop view for hwn and move it to the new desktop
            if (SUCCEEDED(hr)) {
                hr = m_pDesktopManagerInternal->MoveViewToDesktop(pApplicationView, pVirtualDesktop);
                pApplicationView->Release();
            }
            pVirtualDesktop->Release();
            return S_OK;
        }


        if (IsEqualGUID(g, HwndDesktopId))      // Hwnd is already in the good desktop
            return S_OK;


        // hr = this->m_pDesktopManager->MoveWindowToDesktop(hwnd, g);  
        IApplicationView* pApplicationView = nullptr; // Move the hwnd in the desktop owning this case
        hr = this->m_pApplicationViewCollection->GetViewForHwnd(hwnd, &pApplicationView);
        if (SUCCEEDED(hr)) {
            hr = m_pDesktopManagerInternal->FindDesktop(&g, &pVirtualDesktop);
            if (SUCCEEDED(hr)) {
                hr = m_pDesktopManagerInternal->MoveViewToDesktop(pApplicationView, pVirtualDesktop);
                pVirtualDesktop->Release();
            }
            pApplicationView->Release();
        }

        return S_OK;
    }

    // getting FileName for a duplcated handle may hang for event ( console, namedpipe,..). So this operation will be done in another thread
    // if Hang will close the handle - if still hang will terminatethread and create another one 
    WCHAR FileName[MAX_PATH + 1];                                       // used to get DuplicateHandle filename 
    HANDLE hThread;         // handle for the "filename" thread 
    HANDLE hTargetHandleToTranslatedInFileName; // Handle to find the file name - use by the this thread
    HANDLE hEventHandleToBeDecoded; // To communicate with this thread 
    HANDLE hWorkDone;
    volatile BOOL bShutdown;        // to signal whern we will exit - Thread will exit 

    // THread to decode filename 
    static DWORD WINAPI ThreadGetDuplicateHandleFileNameProc(LPVOID lpParameter) {
        LTDesktopManager* pVDC = (LTDesktopManager*)lpParameter;
        do {
            DWORD Status2 = WaitForSingleObject(pVDC->hEventHandleToBeDecoded, INFINITE); // wait for something to do 
            if (WAIT_OBJECT_0 == Status2 && !(pVDC->bShutdown)) {                         // will try to get the filename
                pVDC->FileName[0] = 0;
                DWORD bStatus = GetFinalPathNameByHandleW(pVDC->hTargetHandleToTranslatedInFileName, pVDC->FileName, sizeof(pVDC->FileName) / sizeof(WCHAR) - 1, 0);
                pVDC->FileName[MAX_PATH] = 0;
                // wcout << L"THreadGet " << bStatus << L" " << GetLastError() <<  L" " << pVDC->FileName << endl;
                SetEvent(pVDC->hWorkDone);      // work done 
            }
        } while (!(pVDC->bShutdown));

        return 0;
    }
public:
    LTDesktopManager() : m_CurrentDesktopGuid({ 0 }), m_pServiceProvider(nullptr), m_pDesktopManagerInternal(nullptr), m_pDesktopManager(nullptr), m_pDesktopManagerInternal2(nullptr), m_pApplicationViewCollection(nullptr), bShutdown(false), hThread(nullptr), hEventHandleToBeDecoded(nullptr), hWorkDone(nullptr), hTargetHandleToTranslatedInFileName(nullptr) {
        DWORD dwThreadId = 0;
        hEventHandleToBeDecoded = ::CreateEventW(NULL, false, false, NULL);
        hWorkDone = ::CreateEventW(NULL, false, false, NULL);
        hThread = CreateThread(NULL, NULL, ThreadGetDuplicateHandleFileNameProc, this, 0, &dwThreadId);

    }
    void RemoveAllDesktops() {
        IObjectArray* pObjectArray = nullptr;
        HRESULT hr = m_pDesktopManagerInternal->GetDesktops(&pObjectArray);
        if (SUCCEEDED(hr)) {
            UINT Count;
            do {
                IObjectArray* pObjectArray = nullptr;
                hr = m_pDesktopManagerInternal->GetDesktops(&pObjectArray);
                if (SUCCEEDED(hr)) hr = pObjectArray->GetCount(&Count);
                if (SUCCEEDED(hr)) {
                    if (Count > 1) {
                        IVirtualDesktop* pDesktop1 = nullptr;
                        IVirtualDesktop* pDesktop = nullptr;
                        hr = pObjectArray->GetAt(0, __uuidof(IVirtualDesktop), (void**)&pDesktop1);
                        if (SUCCEEDED(hr)) {
                            hr = pObjectArray->GetAt(Count - 1, __uuidof(IVirtualDesktop), (void**)&pDesktop);
                            if (SUCCEEDED(hr)) {
                                hr = m_pDesktopManagerInternal->RemoveDesktop(pDesktop, pDesktop1);
                            }
                        }
                        if (pDesktop1) pDesktop1->Release();
                        if (pDesktop) pDesktop->Release();
                    }
                    else {
                        hr = E_FAIL;
                        DesktopCase.Reset();
                        this->FindExistingCasesTopDesktop();
                    }
                }
                if( pObjectArray) pObjectArray->Release();
            } while (SUCCEEDED(hr));
            
        }
        return ;
    }
    void DisconnectVirtualDesktopServices() {
        if (m_pServiceProvider) m_pServiceProvider->Release();
        if (m_pDesktopManagerInternal) m_pDesktopManagerInternal->Release();
        if (m_pDesktopManager) m_pDesktopManager->Release();
        if (m_pDesktopManagerInternal2) m_pDesktopManagerInternal2->Release();
        m_pServiceProvider = nullptr;  m_pDesktopManagerInternal = nullptr;  m_pDesktopManager = nullptr; m_pDesktopManagerInternal2 = nullptr;
        memset(&m_CurrentDesktopGuid, 0, sizeof(m_CurrentDesktopGuid));

    }

    void GetMenu( HMENU & hMenu ) {
        UINT Count = 0;
 
        extern wregex CaseId;
        CaseId = L"[1][12][0-9]{13}";
        MENUITEMINFOW mi;

        mi.cbSize = sizeof(MENUITEMINFO);
        mi.fMask = MIIM_TYPE | MIIM_ID | MIIM_STATE;
        mi.fType = MFT_STRING;
        mi.wID = (UINT)SWM_STARTAUTOMATICALLY;
        mi.fState = IsRunAtStartup() ? MFS_CHECKED : MFS_UNCHECKED ;
        mi.dwTypeData = (LPWSTR)L"Start automatically at startup";
        InsertMenuItemW(hMenu, Count++, TRUE, &mi);


        mi.cbSize = sizeof(MENUITEMINFO);
        mi.fMask = MIIM_TYPE;
        mi.fType = MFT_SEPARATOR;
        InsertMenuItemW(hMenu, Count++, TRUE, &mi);


        mi.fMask = MIIM_TYPE | MIIM_ID;
        mi.fType = MFT_STRING;
        mi.wID = (UINT)SWM_ASSIGNUPDATEALLCASESTODESK;
        mi.dwTypeData = (LPWSTR)L"Move/update Cases Windows to Virtual Desktop(s)";
        InsertMenuItemW(hMenu, Count++, TRUE, &mi);

        mi.wID = (UINT)SWM_REMOVEALLDESKTOPS;
        mi.dwTypeData = (LPWSTR)L"Remove all Virtual Desktops";
        InsertMenuItemW(hMenu, Count++, TRUE, &mi);
        
        wstring ClipData = GetClipboardText();
        wsmatch m;

        mi.cbSize = sizeof(MENUITEMINFO);
        mi.fMask = MIIM_TYPE;
        mi.fType = MFT_SEPARATOR;
        InsertMenuItemW(hMenu, Count++, TRUE, &mi);

        regex_search(ClipData, m, CaseId);
        if (m.size()) {
           mi.wID = (UINT)SWM_MOVETOCLIPBOARDESKTOP;
           mi.fMask = MIIM_TYPE | MIIM_ID ;
           wstring s = m[0];
           s = L"Create/update and Move to " + s + L" Virtual Desktop";
           wchar_t* ptr = _wcsdup(s.c_str());  
           mi.dwTypeData = ptr;
           
           InsertMenuItemW(hMenu, Count++, TRUE, &mi);
           free(ptr);
        }
        
        Print(hMenu, &Count); // InsertMenuItem for all desktops
        mi.cbSize = sizeof(MENUITEMINFO);
        mi.fMask = MIIM_TYPE | MIIM_ID;
        mi.wID = (UINT)SWM_EXIT;
        mi.fType = MFT_STRING;
        mi.dwTypeData = (LPWSTR)L"Exit";
        InsertMenuItemW(hMenu, Count++, TRUE, &mi);

    }
    HRESULT ConnectVirtualDesktopServices() {               // Call to connect to varuoius shell interface used to manage virtual desktop
        HRESULT hr;
        hr = ::CoCreateInstance(CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER, __uuidof(IServiceProvider), (PVOID*)&m_pServiceProvider);
        if (SUCCEEDED(hr)) {
            hr = m_pServiceProvider->QueryService(CLSID_VirtualDesktopAPI_Unknown, &m_pDesktopManagerInternal);
        }
        if (SUCCEEDED(hr)) {
            GUID h = __uuidof(IVirtualDesktopManager);
            hr = m_pServiceProvider->QueryService(h, &m_pDesktopManager);
        }
        if (SUCCEEDED(hr)) {
            GUID h = __uuidof(IVirtualDesktopManagerInternal2);
            HRESULT hr2 = m_pServiceProvider->QueryService(h, &m_pDesktopManagerInternal2); // Internal2 not yet implemented 
        }
        if (SUCCEEDED(hr)) {
            HRESULT hr = m_pServiceProvider->QueryService(__uuidof(IApplicationViewCollection), &m_pApplicationViewCollection);
        }
        if (SUCCEEDED(hr)) {
            hr = GetCurrentVirtualDesktop();
        }
        if (FAILED(hr)) {
            DisconnectVirtualDesktopServices();
        }
        return hr;
    }
    ~LTDesktopManager() {           // destructyor - will release shell com interfaces + signal "filename" thread to end
        DisconnectVirtualDesktopServices();
        bShutdown = true;
        SetEvent(this->hEventHandleToBeDecoded);
        WaitForSingleObject(hThread, 500);
        CloseHandle(hThread);
        CloseHandle(hEventHandleToBeDecoded);
        CloseHandle(hWorkDone);
    }
    HRESULT GetCurrentVirtualDesktop()
    {
        ComPtr<IVirtualDesktop> pDesktop;
        HRESULT hr = m_pDesktopManagerInternal->GetCurrentDesktop(&pDesktop);
        if (SUCCEEDED(hr))
        {
            hr = pDesktop->GetID(&m_CurrentDesktopGuid);
            pDesktop->Release();
        }
        return hr;
    }
    void Print(HMENU h = NULL, UINT * Pos = 0 ) {
        // std::wcout << "Main " << GetGuid(m_CurrentDesktopGuid) << std::endl << std::flush;
        DesktopCase.PrintDesktops(h,Pos);
    }

    HRESULT ListViews() {
        IObjectArray* pObjectArray = nullptr;
        HRESULT hr = m_pApplicationViewCollection->GetViews(&pObjectArray);
        if (SUCCEEDED(hr)) {
            UINT count;
            hr = pObjectArray->GetCount(&count);
            if (SUCCEEDED(hr)) {
                wcout << L"View count " << count << endl;
                for (UINT i = 0; i < count; i++) {
                    IWin32ApplicationView* pView = nullptr;
                    IInspectable* pInspectable = nullptr;

                    // hr = pObjectArray->GetAt(i, __uuidof(IInspectable), (void**)&pInspectable);
                    // ULONG iidcount = 0;
                    // IID* p;
                    // hr = pInspectable->GetIids(&iidcount,&p );
                    // for (ULONG  k = 0; k < iidcount; k++) {
                    //    wcout << L"Iiid " << GetGuid(p[k]) << endl ;
                    // }

                    hr = pObjectArray->GetAt(i, __uuidof(IWin32ApplicationView), (void**)&pView);
                    if (FAILED(hr)) continue;
                    HWND hwnd = 0;
                    pView->GetWindow(&hwnd);

                    wcout << "View " << dec << hwnd << endl;
                    pView->Release();
                }
            }
        }
        return hr;
    }
    HRESULT ListVirtualDesktops() {
        IObjectArray* pObjectArray = nullptr;
        HRESULT hr = m_pDesktopManagerInternal->GetDesktops(&pObjectArray);

        if (SUCCEEDED(hr)) {
            UINT count;
            hr = pObjectArray->GetCount(&count);
            if (SUCCEEDED(hr)) {
                for (UINT i = 0; i < count; i++) {
                    IVirtualDesktop* pDesktop = nullptr;
                    if (FAILED(pObjectArray->GetAt(i, __uuidof(IVirtualDesktop), (void**)&pDesktop)))   continue;

                    GUID id = { 0 };
                    if (SUCCEEDED(pDesktop->GetID(&id))) {
                        // std::wcout << L"\t #" << i << L": ";
                        // std::wcout << GetGuid(id);
                        // std::wcout << std::endl;
                        if (i == 0) {
                            m_CurrentDesktopGuid = id;
                        }
                        this->DesktopCase.AddDesktop(id);
                    }
                    pDesktop->Release();
                }
            }
            pObjectArray->Release();
        }
        return hr;
    }

    static BOOL CALLBACK   EnumFirstPass(_In_ HWND   hwnd, _In_ LPARAM lParam) {
        WCHAR Title[500];
        LTDesktopManager* pvt = (LTDesktopManager*)lParam;

        int i = ::GetWindowTextW(hwnd, Title, (sizeof(Title) / sizeof(Title[0])) - 1);
        Title[i] = L'\0';
        wstring wttile(Title);
        wsmatch m;
        wregex AllCaseId(L"[1][12][0-9]{13}");
        regex_search(wttile, m, AllCaseId);
        if (m.size()) {
            GUID desktopId;
            HRESULT hr = (pvt->m_pDesktopManager)->GetWindowDesktopId(hwnd, &desktopId);

            if (SUCCEEDED(hr)) {
                // wcout << "*** EnumFirstPass hwnd=" << hwnd << " in desktop " << GetGuid(desktopId) << L" / " << Title << std::endl;
                if (!IsEqualGUID(desktopId, pvt->m_CurrentDesktopGuid)) {
                    GUID g = pvt->DesktopCase.GetDesktopId(m[0]);
                    if (IsEqualGUID(g, { 0 })) {
                        // wcout << "EnumFirstPass  caseid key " << m[0] << " not yet in map" << endl;
                        BOOL b = pvt->DesktopCase.IsDesktopIdAlreadyInUse(desktopId);
                        if (!b)  pvt->DesktopCase.AddDesktopId(m[0], desktopId);
                        else {
                            // wcout << "EnumFirstPass  Caseid key " << m[0] << " not added - DesktopId already used " << endl;
                        }
                    }
                }
            }
            DWORD dwProcessId = 0;
            map<DWORD, DWORD>& mProcessAlreadyProcessed = pvt->m_ProcessAlreadyProcessed;
            DWORD dwThread = GetWindowThreadProcessId(hwnd, &dwProcessId);

            auto ProcessIdInMap = mProcessAlreadyProcessed.find(dwProcessId);
            if (ProcessIdInMap == mProcessAlreadyProcessed.end())  mProcessAlreadyProcessed[dwProcessId] = 1;

        }
        else {

        }
        return true;

    }

    HRESULT SwitchToVirtualDesktop(int DesktopNumber) {
        GUID g;
        HRESULT hr = S_OK;
        g = this->DesktopCase.GetDesktopGuid(DesktopNumber);

        if (!IsEqualGUID({ 0 }, g)) {
            IVirtualDesktop* pVirtualdesktop = nullptr;
            hr = m_pDesktopManagerInternal->FindDesktop(&g, &pVirtualdesktop);
            if (SUCCEEDED(hr)) {
                m_pDesktopManagerInternal->SwitchDesktop(pVirtualdesktop);
                pVirtualdesktop->Release();
            }
        }
        return hr;
    }
    HRESULT SwitchToVirtualDesktop(wstring CaseToSwitch) {
        GUID g;
        HRESULT hr = S_OK;
        g = this->DesktopCase.GetDesktopId(CaseToSwitch);

        if (!IsEqualGUID({ 0 }, g)) {
            IVirtualDesktop* pVirtualdesktop = nullptr;
            hr = m_pDesktopManagerInternal->FindDesktop(&g, &pVirtualdesktop);
            if (SUCCEEDED(hr)) {
                m_pDesktopManagerInternal->SwitchDesktop(pVirtualdesktop);
                pVirtualdesktop->Release();
            }
        }
        return hr;
    }
    static BOOL CALLBACK   EnumSecondPass(_In_ HWND   hwnd, _In_ LPARAM lParam) {
        WCHAR Title[500];
        LTDesktopManager* pvt = (LTDesktopManager*)lParam;
        map<DWORD, DWORD>& mProcessAlreadyProcessed = pvt->m_ProcessAlreadyProcessed;

        int i = ::GetWindowTextW(hwnd, Title, (sizeof(Title) / sizeof(Title[0])) - 1);
        Title[i] = L'\0';
        wstring wttile(Title);
        wsmatch m;

        DWORD dwProcessId = 0;
        DWORD dwThread = GetWindowThreadProcessId(hwnd, &dwProcessId);

        regex_search(wttile, m, CaseId);

        if (m.size()) {
            GUID desktopId;
            HRESULT hr = (pvt->m_pDesktopManager)->GetWindowDesktopId(hwnd, &desktopId);

            if (SUCCEEDED(hr)) {
                // wcout << "*** EnumAndMoveWindowsProc hwnd=" << hwnd << " in desktop " << GetGuid(desktopId) << L" / " << Title << std::endl;
                pvt->MoveToDesktop(hwnd, desktopId, m[0]);
            }
            if (m.size()) mProcessAlreadyProcessed[dwProcessId] = 1;
            return true;
        }

        auto ProcessIdInMap = mProcessAlreadyProcessed.find(dwProcessId);
        BOOL bProcessFirst = (ProcessIdInMap == mProcessAlreadyProcessed.end());

        if (bProcessFirst) {
            mProcessAlreadyProcessed[dwProcessId] = 0;
            DWORD  ProcId = dwProcessId;

            HANDLE hProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, ProcId);
            if (INVALID_HANDLE_VALUE != hProc) {
                WCHAR ProcImage[MAX_PATH + 1];
                DWORD dwLen = GetProcessImageFileNameW(hProc, ProcImage, sizeof(ProcImage) / sizeof(WCHAR) - 1);
                ProcImage[dwLen] = 0;
                // wcout << ProcImage;
                CloseHandle(hProc);
                DWORD ProcessHandleTable = 58;
                hProc = OpenProcess(PROCESS_ALL_ACCESS, FALSE, ProcId);

                if (INVALID_HANDLE_VALUE != hProc) {
                    HPSS SnapshotHandle;
                    DWORD dwSnap = PssCaptureSnapshot(hProc, PSS_CAPTURE_HANDLES | PSS_CAPTURE_HANDLE_NAME_INFORMATION | PSS_CAPTURE_HANDLE_TYPE_SPECIFIC_INFORMATION, NULL, &SnapshotHandle);
                    if (ERROR_SUCCESS == dwSnap) {
                        DWORD dwHandleCount = 0;
                        DWORD dwBufferLength = sizeof(dwHandleCount);
                        DWORD status = PssQuerySnapshot(SnapshotHandle, PSS_QUERY_HANDLE_INFORMATION, &dwHandleCount, dwBufferLength);
                        HPSSWALK WalkMarkerHandle;
                        PssWalkMarkerCreate(nullptr, &WalkMarkerHandle);
                        ULONG c = 0;
                        for (; c < dwHandleCount; c++) {
                            PSS_HANDLE_ENTRY Handle;
                            DWORD rc = PssWalkSnapshot(SnapshotHandle, PSS_WALK_HANDLES, WalkMarkerHandle, &Handle, sizeof(Handle));
                            if (Handle.ObjectType == PSS_OBJECT_TYPE_UNKNOWN) { // File 
                                HANDLE hTargetHandle = nullptr;
                                BOOL bStatus = DuplicateHandle(hProc, Handle.Handle, GetCurrentProcess(), &hTargetHandle, DUPLICATE_SAME_ACCESS, FALSE, DUPLICATE_SAME_ACCESS);
                                if (bStatus) {
                                    pvt->hTargetHandleToTranslatedInFileName = hTargetHandle;
                                    // wcout << ProcImage << L" " << Handle.Handle << L" " << hTargetHandle << endl << flush;

                                    SetEvent(pvt->hEventHandleToBeDecoded);
                                    DWORD dw = WaitForSingleObject(pvt->hWorkDone, 100);
                                    if (dw == WAIT_TIMEOUT) {
                                        // wcout << L"Tiemout\n" << endl;
                                        CloseHandle(hTargetHandle); hTargetHandle = nullptr;
                                        dw = WaitForSingleObject(pvt->hWorkDone, 30);
                                        if (dw == WAIT_TIMEOUT) {
                                            // wcout << L"TerminateThread\n" << endl;
                                            TerminateThread(pvt->hThread, 0);
                                            ResetEvent(pvt->hEventHandleToBeDecoded);
                                            ResetEvent(pvt->hWorkDone);
                                            CloseHandle(pvt->hThread);
                                            DWORD dwThreadId = 0;
                                            pvt->hThread = CreateThread(NULL, NULL, ThreadGetDuplicateHandleFileNameProc, pvt, 0, &dwThreadId);
                                        }

                                    }
                                    wstring obj(pvt->FileName);
                                    regex_search(obj, m, CaseId);
                                    if (m.size()) {
                                        // wcout << L"Process  " << obj << L"/Pid " << ProcId << L"/ ProcImage" << ProcImage  << endl;
                                        mProcessAlreadyProcessed[dwProcessId] = 2;
                                        if (hTargetHandle) CloseHandle(hTargetHandle);
                                        hTargetHandle = nullptr;
                                        GUID desktopId;
                                        HRESULT hr = (pvt->m_pDesktopManager)->GetWindowDesktopId(hwnd, &desktopId);

                                        if (SUCCEEDED(hr)) {
                                            // wcout << "*** EnumAndMoveWindowsProc hwnd=" << hwnd << " in desktop " << GetGuid(desktopId) << L" / " << Title << std::endl;
                                            pvt->MoveToDesktop(hwnd, desktopId, m[0]);
                                        }

                                        break;
                                    }
                                }
                                if (hTargetHandle) CloseHandle(hTargetHandle);
                            }

                        }
                        PssWalkMarkerFree(WalkMarkerHandle);
                        PssFreeSnapshot(hProc, SnapshotHandle);
                    }
                    CloseHandle(hProc);
                }
            }
        }

        regex_search(wttile, m, CaseId);


        return true;
    }
    void FindExistingCasesTopDesktop() {
        ListVirtualDesktops();
        m_ProcessAlreadyProcessed.clear();
        EnumWindows(EnumFirstPass, (LPARAM)this);

    }
    void MoveCasesTopWindows() {
        EnumWindows(EnumFirstPass, (LPARAM)this);
        EnumWindows(EnumSecondPass, (LPARAM)this);

    }
};


wregex Move(L"[/-](m|mo|mov|move|movep|movepi|movepid)[:=]([0-9][0-9]*)");
wregex SwitchToDesktop(L"[/-](s|sw|swi|swit|switc|switch)");

LTDesktopManager vtGUI ;


HRESULT InitDesktopManagerGUI(BOOL bDebug ) {
    if (bDebug) {
        freopen("CONOUT$", "w", stdout);

        freopen("CONIN$", "r", stdin);
        SetConsoleTitle(L"Debug Console");
        std::wcout.clear();
        std::wcin.clear();
        std::ios::sync_with_stdio(1);
    };

    ::CoInitialize(NULL);
    HRESULT hr = vtGUI.ConnectVirtualDesktopServices();
    if (FAILED(hr)) {
        return hr;
    }
    vtGUI.FindExistingCasesTopDesktop();
    wcout << "Existing Desktop list\n" << endl;
    vtGUI.Print();

    return hr;
}


int main(int argc, TCHAR* argv[])
{
    ::CoInitialize(NULL);
    LTDesktopManager vt;
    wstring CaseTowork;
    boolean bSwitchToDesktop = false;
    boolean bMovePid = false;
    DWORD dwPid = 0;
    HRESULT hr = vt.ConnectVirtualDesktopServices();

    WCHAR* CMdLine = GetCommandLineW();
    wsmatch m;
    wstring cmdline(CMdLine);
    regex_search(cmdline, m, CaseId);
    if (m.size()) {
        CaseTowork = m[0];
        wcout << L"With case " << CaseTowork << endl;
    }
    regex_search(cmdline, m, SwitchToDesktop);
    if (m.size()) {
        bSwitchToDesktop = true;

    }
    regex_search(cmdline, m, Move);
    if (m.size() == 3) {
        bMovePid = true;
        dwPid = std::stol(m[2]);
        wcout << L"Will move PID " << dwPid << " windows to VirtualDesktop " << CaseTowork << endl;
    }
    if (SUCCEEDED(hr)) {
        vt.ListVirtualDesktops();

        vt.Print();
        wcout << L" " << endl;
        // vt.ListViews();
        vt.MoveCasesTopWindows();
        vt.Print();
        

    }
    if (bSwitchToDesktop) {
        wcout << "Will switch to desktop hosting " << CaseTowork << endl;
        vt.SwitchToVirtualDesktop(CaseTowork);
    }
    vt.DisconnectVirtualDesktopServices();
}


#pragma region Tray

#define TRAYICONID	1//				ID number for the Notify Icon
#define SWM_TRAYMSG	WM_APP//		the message ID sent to our window




// Global Variables:
HINSTANCE		hInst;	// current instance
NOTIFYICONDATA	niData;	// notify icon data

// Forward declarations of functions included in this code module:
BOOL				InitInstance(HINSTANCE, int);
BOOL				OnInitDialog(HWND hWnd);
void				ShowContextMenu(HWND hWnd);
ULONGLONG			GetDllVersion(LPCTSTR lpszDllName);

INT_PTR CALLBACK	DlgProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK	About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY _tWinMain(HINSTANCE hInstance,
                     HINSTANCE hPrevInstance,
                     LPTSTR    lpCmdLine,
                     int       nCmdShow)
{
    AllocConsole();


	MSG msg;
	HACCEL hAccelTable;


    InitDesktopManagerGUI(TRUE);



	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow)) return FALSE;
	hAccelTable = LoadAccelerators(hInstance, (LPCTSTR)IDC_STEALTHDIALOG);

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg)||
			!IsDialogMessage(msg.hwnd,&msg) ) 
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	return (int) msg.wParam;
}

//	Initialize the window and tray icon
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	// prepare for XP style controls
	InitCommonControls();

	 // store instance handle and create dialog
	hInst = hInstance;
	HWND hWnd = CreateDialog( hInstance, MAKEINTRESOURCE(IDD_DLG_DIALOG),
		NULL, (DLGPROC)DlgProc );
	if (!hWnd) return FALSE;

	// Fill the NOTIFYICONDATA structure and call Shell_NotifyIcon

	// zero the structure - note:	Some Windows funtions require this but
	//								I can't be bothered which ones do and
	//								which ones don't.
	ZeroMemory(&niData,sizeof(NOTIFYICONDATA));

	// get Shell32 version number and set the size of the structure
	//		note:	the MSDN documentation about this is a little
	//				dubious and I'm not at all sure if the method
	//				bellow is correct
	ULONGLONG ullVersion = GetDllVersion(_T("Shell32.dll"));
	if(ullVersion >= MAKEDLLVERULL(5, 0,0,0))
		niData.cbSize = sizeof(NOTIFYICONDATA);
	else niData.cbSize = NOTIFYICONDATA_V2_SIZE;

	// the ID number can be anything you choose
	niData.uID = TRAYICONID;

	// state which structure members are valid
	niData.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;

	// load the icon
	niData.hIcon = (HICON)LoadImage(hInstance,MAKEINTRESOURCE(IDI_STEALTHDLG),
		IMAGE_ICON, GetSystemMetrics(SM_CXSMICON),GetSystemMetrics(SM_CYSMICON),
		LR_DEFAULTCOLOR);

	// the window to send messages to and the message to send
	//		note:	the message value should be in the
	//				range of WM_APP through 0xBFFF
	niData.hWnd = hWnd;
    niData.uCallbackMessage = SWM_TRAYMSG;

	// tooltip message
    lstrcpyn(niData.szTip, _T("Place Cases related windows to a\ndedicated Virtual Desktop\nAllow to switch to these case related virtual desktops"), sizeof(niData.szTip)/sizeof(TCHAR));

	Shell_NotifyIcon(NIM_ADD,&niData);

	// free icon handle
	if(niData.hIcon && DestroyIcon(niData.hIcon))
		niData.hIcon = NULL;

	// call ShowWindow here to make the dialog initially visible

	return TRUE;
}

BOOL OnInitDialog(HWND hWnd)
{
	HMENU hMenu = GetSystemMenu(hWnd,FALSE);
	if (hMenu)
	{
		AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
		AppendMenu(hMenu, MF_STRING, IDM_ABOUT, _T("About"));
	}
	HICON hIcon = (HICON)LoadImage(hInst,
		MAKEINTRESOURCE(IDI_STEALTHDLG),
		IMAGE_ICON, 0,0, LR_SHARED|LR_DEFAULTSIZE);
	SendMessage(hWnd,WM_SETICON,ICON_BIG,(LPARAM)hIcon);
	SendMessage(hWnd,WM_SETICON,ICON_SMALL,(LPARAM)hIcon);
	return TRUE;
}

HMENU hMenu = NULL;
// Name says it all
void ShowContextMenu(HWND hWnd)
{


	POINT pt;
	GetCursorPos(&pt);
	hMenu = CreatePopupMenu();
	if(hMenu)
	{
        vtGUI.GetMenu(hMenu);

		// note:	must set window to the foreground or the
		//			menu won't disappear when it should
		SetForegroundWindow(hWnd);

		TrackPopupMenu(hMenu, TPM_BOTTOMALIGN,
			pt.x, pt.y, 0, hWnd, NULL );

	}
}

// Get dll version number
ULONGLONG GetDllVersion(LPCTSTR lpszDllName)
{
    ULONGLONG ullVersion = 0;
	HINSTANCE hinstDll;
    hinstDll = LoadLibrary(lpszDllName);
    if(hinstDll)
    {
        DLLGETVERSIONPROC pDllGetVersion;
        pDllGetVersion = (DLLGETVERSIONPROC)GetProcAddress(hinstDll, "DllGetVersion");
        if(pDllGetVersion)
        {
            DLLVERSIONINFO dvi;
            HRESULT hr;
            ZeroMemory(&dvi, sizeof(dvi));
            dvi.cbSize = sizeof(dvi);
            hr = (*pDllGetVersion)(&dvi);
            if(SUCCEEDED(hr))
				ullVersion = MAKEDLLVERULL(dvi.dwMajorVersion, dvi.dwMinorVersion,0,0);
        }
        FreeLibrary(hinstDll);
    }
    return ullVersion;
}

// Message handler for the app
INT_PTR CALLBACK DlgProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;

	switch (message) 
	{
	case SWM_TRAYMSG:
		switch(lParam)
		{
		case WM_LBUTTONDBLCLK:
			ShowWindow(hWnd, SW_RESTORE);
			break;
		case WM_RBUTTONDOWN:
		case WM_CONTEXTMENU:
			ShowContextMenu(hWnd);
		}
		break;
	case WM_SYSCOMMAND:
		if((wParam & 0xFFF0) == SC_MINIMIZE)
		{
			ShowWindow(hWnd, SW_HIDE);
			return 1;
		}
		else if(wParam == IDM_ABOUT)
			DialogBox(hInst, (LPCTSTR)IDD_ABOUTBOX, hWnd, (DLGPROC)About);
		break;
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam); 

		switch (wmId)
		{
        case SWM_STARTAUTOMATICALLY: 
                if (IsRunAtStartup() )
                    IsRunAtStartup(TRUE, FALSE);
                else
                    IsRunAtStartup(FALSE, TRUE);
                
            break;
		case SWM_EXIT:
			DestroyWindow(hWnd);
			break;
        case SWM_REMOVEALLDESKTOPS :
            vtGUI.RemoveAllDesktops();
            break;
        case SWM_ASSIGNUPDATEALLCASESTODESK:
            extern wregex CaseId;
            CaseId = L"[1][12][0-9]{13}";
            vtGUI.MoveCasesTopWindows();
            vtGUI.Print();
            break;
        default:
        // case SWM_MOVETOCLIPBOARDESKTOP:
  
            if (SWM_MOVETOCLIPBOARDESKTOP == wmId ) {        
                MENUITEMINFOW mi;
                mi.cbSize = sizeof(MENUITEMINFOW);
                CaseId = L"[1][12][0-9]{13}";
                mi.fMask = MIIM_STRING;
                wchar_t strText[60];
                mi.dwTypeData = strText;
                mi.cch = ARRAYSIZE(strText);
                BOOL Success = GetMenuItemInfoW(hMenu, wmId, FALSE, &mi);
                std::wcout << L"Will execute menuitem " << strText << L" " << hex << wmId << dec << endl;              wsmatch m;
                wstring s(strText);
                regex_search(s, m, CaseId);
                if (m.size() > 0) {
                    CaseId = wstring(m[0]);
                    vtGUI.MoveCasesTopWindows();
                    vtGUI.FindExistingCasesTopDesktop();
                    vtGUI.Print();
                }
                vtGUI.SwitchToVirtualDesktop(wstring(m[0]));
            }
            else {
                int wDesktop = wmId - (SWM_MOVETODESKTOP);
                vtGUI.SwitchToVirtualDesktop(wDesktop);

            }
            break;
   
            break;


		}
        if (hMenu != NULL) {
            DestroyMenu(hMenu);
            hMenu = NULL;
        }
		return 1;
	case WM_INITDIALOG:
		return OnInitDialog(hWnd);
	case WM_CLOSE:
		DestroyWindow(hWnd);
		break;
	case WM_DESTROY:
		niData.uFlags = 0;
		Shell_NotifyIcon(NIM_DELETE,&niData);
		PostQuitMessage(0);
		break;
	}
	return 0;
}

// Message handler for about box.
LRESULT CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_INITDIALOG:
		return TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) 
		{
			EndDialog(hDlg, LOWORD(wParam));
			return TRUE;
		}
		break;
	}
	return FALSE;
}

#pragma endregion Tray