#include <windows.h>
#include <tchar.h>
#include <shlwapi.h>
#define TVTEST_PLUGIN_CLASS_IMPLEMENT
#include "TVTestPlugin.h"
#include "TVTestInterface.h"
#include "TVCAS.h"
#include "CasProcessor.h"
#include "resource.h"

#pragma comment(lib, "shlwapi.lib")




class CLock
{
public:
	CLock() { ::InitializeCriticalSection(&m_CriticalSection); }
	~CLock() { ::DeleteCriticalSection(&m_CriticalSection); }
	void Lock() { ::EnterCriticalSection(&m_CriticalSection); }
	void Unlock() { ::LeaveCriticalSection(&m_CriticalSection); }

private:
	CRITICAL_SECTION m_CriticalSection;
};

class CBlockLock
{
public:
	CBlockLock(CLock &Lock) : m_Lock(Lock) { m_Lock.Lock(); }
	~CBlockLock() { m_Lock.Unlock(); }

private:
	CLock &m_Lock;
};

class CIUnknownImpl
{
public:
	CIUnknownImpl() : m_RefCount(1) {}

	ULONG AddRefImpl() {
		return ::InterlockedIncrement(&m_RefCount);
	}

	ULONG ReleaseImpl() {
		LONG Count = ::InterlockedDecrement(&m_RefCount);
		if (Count == 0)
			delete this;
		return Count;
	}

protected:
	LONG m_RefCount;

	virtual ~CIUnknownImpl() {}
};




static LPOLESTR AllocateOleString(LPCOLESTR pszSrc)
{
	size_t Size = (::lstrlenW(pszSrc) + 1) * sizeof(WCHAR);
	LPOLESTR pszDst = static_cast<LPOLESTR>(::CoTaskMemAlloc(Size));
	if (pszDst != nullptr)
		::CopyMemory(pszDst, pszSrc, Size);
	return pszDst;
}


static void StringFormatAppend(LPTSTR pszString, int Length, int *pPos, LPCTSTR pszFormat, ...)
{
	int Pos = *pPos;

	if (Pos >= Length - 1)
		return;

	va_list Args;

	va_start(Args, pszFormat);
	int Result = ::wvnsprintf(pszString + Pos, Length - Pos, pszFormat, Args);
	va_end(Args);
	if (Result <= 0)
		pszString[Pos] = '\0';
	else
		*pPos = Pos + Result;
}




class CEnumFilterModule
	: public TVTest::Interface::IEnumFilterModule
	, protected CIUnknownImpl
{
public:
	CEnumFilterModule();

// IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject) override;
	STDMETHODIMP_(ULONG) AddRef() override { return AddRefImpl(); }
	STDMETHODIMP_(ULONG) Release() override { return ReleaseImpl(); }

// IEnumFilterModule
	STDMETHODIMP Next(BSTR *pName) override;

private:
	~CEnumFilterModule();

	HANDLE m_hFind;
	WIN32_FIND_DATAW m_FindData;
	bool m_fFirst;
	bool m_fNext;
};


CEnumFilterModule::CEnumFilterModule()
	: m_hFind(INVALID_HANDLE_VALUE)
	, m_fFirst(true)
	, m_fNext(false)
{
}


CEnumFilterModule::~CEnumFilterModule()
{
	if (m_hFind != INVALID_HANDLE_VALUE)
		::FindClose(m_hFind);
}


STDMETHODIMP CEnumFilterModule::QueryInterface(REFIID riid, void **ppvObject)
{
	if (ppvObject == nullptr)
		return E_POINTER;

	if (riid == IID_IUnknown || riid == __uuidof(IEnumFilterModule)) {
		*ppvObject = static_cast<IEnumFilterModule*>(this);
	} else {
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	AddRef();

	return S_OK;
}


STDMETHODIMP CEnumFilterModule::Next(BSTR *pName)
{
	if (pName == nullptr)
		return E_POINTER;

	if (m_fFirst) {
		WCHAR szPath[MAX_PATH];
		DWORD Result = ::GetModuleFileNameW(nullptr, szPath, _countof(szPath));

		if (Result > 0 && Result < _countof(szPath)) {
			::PathRemoveFileSpecW(szPath);
			::PathAppendW(szPath, L"*.tvcas");
			m_hFind = ::FindFirstFileW(szPath, &m_FindData);
		}
		m_fFirst = false;
	}

	if (m_hFind == INVALID_HANDLE_VALUE)
		return S_FALSE;

	while (m_fNext || (m_FindData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
		if (!::FindNextFileW(m_hFind, &m_FindData)) {
			::FindClose(m_hFind);
			m_hFind = INVALID_HANDLE_VALUE;
			return S_FALSE;
		}
		m_fNext = false;
	}
	m_fNext = true;

	*pName = ::SysAllocString(m_FindData.cFileName);
	if (*pName == nullptr)
		return E_OUTOFMEMORY;

	return S_OK;
}




class CFilterDevice
	: public TVTest::Interface::IFilterDevice
	, protected CIUnknownImpl
{
public:
	CFilterDevice(TVCAS::ICasDevice *pCasDevice);

// IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject) override;
	STDMETHODIMP_(ULONG) AddRef() override { return AddRefImpl(); }
	STDMETHODIMP_(ULONG) Release() override { return ReleaseImpl(); }

// IFilterDevice
	STDMETHODIMP GetDeviceInfo(TVTest::Interface::FilterDeviceInfo *pInfo) override;
	STDMETHODIMP GetFilterCount(int *pCount) override;
	STDMETHODIMP GetFilterName(int Index, BSTR *pName) override;
	STDMETHODIMP IsFilterAvailable(LPCWSTR pszName, BOOL *pfAvailable) override;

private:
	~CFilterDevice();

	TVCAS::ICasDevice *m_pCasDevice;
};


CFilterDevice::CFilterDevice(TVCAS::ICasDevice *pCasDevice)
	: m_pCasDevice(pCasDevice)
{
}


CFilterDevice::~CFilterDevice()
{
	m_pCasDevice->Release();
}


STDMETHODIMP CFilterDevice::QueryInterface(REFIID riid, void **ppvObject)
{
	if (ppvObject == nullptr)
		return E_POINTER;

	if (riid == IID_IUnknown || riid == __uuidof(IFilterDevice)) {
		*ppvObject = static_cast<IFilterDevice*>(this);
	} else {
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	AddRef();

	return S_OK;
}


STDMETHODIMP CFilterDevice::GetDeviceInfo(TVTest::Interface::FilterDeviceInfo *pInfo)
{
	if (pInfo == nullptr)
		return E_POINTER;

	TVCAS::CasDeviceInfo DeviceInfo;

	if (!m_pCasDevice->GetDeviceInfo(&DeviceInfo)) {
		::ZeroMemory(pInfo, sizeof(TVTest::Interface::FilterDeviceInfo));
		return E_FAIL;
	}

	pInfo->DeviceID = DeviceInfo.DeviceID;
	pInfo->Flags = DeviceInfo.Flags;
	pInfo->Name = ::SysAllocString(DeviceInfo.Name);
	pInfo->Text = ::SysAllocString(DeviceInfo.Text);

	return S_OK;
}


STDMETHODIMP CFilterDevice::GetFilterCount(int *pCount)
{
	if (pCount == nullptr)
		return E_POINTER;

	*pCount = m_pCasDevice->GetCardCount();

	return S_OK;
}


STDMETHODIMP CFilterDevice::GetFilterName(int Index, BSTR *pName)
{
	if (pName == nullptr)
		return E_POINTER;

	WCHAR szName[256];

	if (!m_pCasDevice->GetCardName(Index, szName, _countof(szName))) {
		*pName = nullptr;
		return E_FAIL;
	}

	*pName = ::SysAllocString(szName);
	if (*pName == nullptr)
		return E_OUTOFMEMORY;

	return S_OK;
}


STDMETHODIMP CFilterDevice::IsFilterAvailable(LPCWSTR pszName, BOOL *pfAvailable)
{
	if (pfAvailable == nullptr)
		return E_POINTER;

	*pfAvailable = m_pCasDevice->IsCardAvailable(pszName);

	return S_OK;
}




class CFilter
	: public TVTest::Interface::IFilter
	, protected CIUnknownImpl
{
public:
	CFilter(TVCAS::ICasManager *pCasManager, int Device);

// IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject) override;
	STDMETHODIMP_(ULONG) AddRef() override { return AddRefImpl(); }
	STDMETHODIMP_(ULONG) Release() override { return ReleaseImpl(); }

// IFilter
	STDMETHODIMP GetDeviceID(TVTest::Interface::DeviceIDType *pDeviceID) override;
	STDMETHODIMP GetName(BSTR *pName) override;
	STDMETHODIMP SendCommand(const void *pSendData, DWORD SendSize, void *pRecvData, DWORD *pRecvSize) override;

private:
	~CFilter();

	TVCAS::ICasManager *m_pCasManager;
	int m_Device;
};


CFilter::CFilter(TVCAS::ICasManager *pCasManager, int Device)
	: m_pCasManager(pCasManager)
	, m_Device(Device)
{
	m_pCasManager->Refer();
}


CFilter::~CFilter()
{
	m_pCasManager->Release();
}


STDMETHODIMP CFilter::QueryInterface(REFIID riid, void **ppvObject)
{
	if (ppvObject == nullptr)
		return E_POINTER;

	if (riid == IID_IUnknown || riid == __uuidof(IFilter)) {
		*ppvObject = static_cast<IFilter*>(this);
	} else {
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	AddRef();

	return S_OK;
}


STDMETHODIMP CFilter::GetDeviceID(TVTest::Interface::DeviceIDType *pDeviceID)
{
	if (pDeviceID == nullptr)
		return E_POINTER;

	TVCAS::CasDeviceInfo DeviceInfo;

	if (!m_pCasManager->GetCasDeviceInfo(m_Device, &DeviceInfo)) {
		*pDeviceID = 0;
		return E_FAIL;
	}

	*pDeviceID = DeviceInfo.DeviceID;

	return S_OK;
}


STDMETHODIMP CFilter::GetName(BSTR *pName)
{
	if (pName == nullptr)
		return E_POINTER;

	WCHAR szName[256];

	if (m_pCasManager->GetCasCardName(szName, _countof(szName)) < 1) {
		*pName = nullptr;
		return E_FAIL;
	}

	*pName = ::SysAllocString(szName);
	if (*pName == nullptr)
		return E_OUTOFMEMORY;

	return S_OK;
}


STDMETHODIMP CFilter::SendCommand(const void *pSendData, DWORD SendSize, void *pRecvData, DWORD *pRecvSize)
{
	if (!m_pCasManager->SendCasCommand(pSendData, SendSize, pRecvData, pRecvSize))
		return E_FAIL;
	return S_OK;
}




class CFilterModuleImpl
	: public TVTest::Interface::IFilterModule
{
public:
	CFilterModuleImpl();

// IFilterModule
	STDMETHODIMP LoadModule(LPCWSTR pszName) override;
	STDMETHODIMP UnloadModule() override;
	STDMETHODIMP IsModuleLoaded() override;
	STDMETHODIMP GetModuleInfo(TVTest::Interface::FilterModuleInfo *pInfo) override;
	STDMETHODIMP GetDeviceCount(int *pCount) override;
	STDMETHODIMP GetDevice(int Device, TVTest::Interface::IFilterDevice **ppDevice) override;
	STDMETHODIMP IsDeviceAvailable(int Device, BOOL *pfAvailable) override;
	STDMETHODIMP CheckDeviceAvailability(int Device, BOOL *pfAvailable, BSTR *pMessage) override;
	STDMETHODIMP GetDefaultDevice(int *pDevice) override;
	STDMETHODIMP GetDeviceByID(TVTest::Interface::DeviceIDType DeviceID, int *pDevice) override;
	STDMETHODIMP GetDeviceByName(LPCWSTR pszName, int *pDevice) override;
	STDMETHODIMP OpenFilter(int Device, LPCWSTR pszName, TVTest::Interface::IFilter **ppFilter) override;

protected:
	HMODULE m_hModule;
	TVCAS::ICasManager *m_pCasManager;
	TVCAS::ModuleInfo m_CasModuleInfo;

	~CFilterModuleImpl();
	virtual HRESULT OnError(HRESULT hr, LPCWSTR pszText, LPCWSTR pszAdvise = nullptr, LPCWSTR pszSystemMessage = nullptr) { return E_NOTIMPL;}
};


CFilterModuleImpl::CFilterModuleImpl()
	: m_hModule(nullptr)
	, m_pCasManager(nullptr)
{
}


CFilterModuleImpl::~CFilterModuleImpl()
{
	UnloadModule();
}


STDMETHODIMP CFilterModuleImpl::LoadModule(LPCWSTR pszName)
{
	if (pszName == nullptr)
		return E_POINTER;

	if (m_hModule != nullptr)
		return E_UNEXPECTED;

	m_hModule = ::LoadLibraryW(pszName);
	if (m_hModule == nullptr) {
		const DWORD ErrorCode = ::GetLastError();
		HRESULT hr = HRESULT_FROM_WIN32(ErrorCode);
		WCHAR szText[MAX_PATH + 32], szAdvise[256];

		::wnsprintfW(szText, _countof(szText),
			L"CASライブラリ \"%s\" をロードできません。", pszName);

		switch (ErrorCode) {
		case ERROR_MOD_NOT_FOUND:
			::lstrcpynW(szAdvise, L"ファイルが見つかりません。", _countof(szAdvise));
			break;

		case ERROR_BAD_EXE_FORMAT:
			::lstrcpynW(szAdvise,
#ifndef _WIN64
				L"32"
#else
				L"64"
#endif
				L"ビット用のDLLではないか、ファイルが破損している可能性があります。",
				_countof(szAdvise));
			break;

		default:
			::wnsprintfW(szAdvise, _countof(szText), L"エラーコード: 0x%x", ErrorCode);
		}

		OnError(hr, szText, szAdvise);

		return hr;
	}

	TVCAS::GetModuleInfoFunc pGetModuleInfo = TVCAS::Helper::Module::GetModuleInfo(m_hModule);
	if (pGetModuleInfo == nullptr
			|| !pGetModuleInfo(&m_CasModuleInfo)
			|| m_CasModuleInfo.LibVersion != TVCAS::LIB_VERSION) {
		UnloadModule();
		OnError(
			E_FAIL,
			pGetModuleInfo == nullptr ?
				L"指定されたDLLがCASライブラリではありません。" :
				L"CASライブラリのバージョンが非対応です。");
		return E_FAIL;
	}

	TVCAS::CreateInstanceFunc pCreateInstance = TVCAS::Helper::Module::CreateInstance(m_hModule);
	if (pCreateInstance == nullptr) {
		UnloadModule();
		OnError(E_FAIL, L"CASライブラリから必要な関数を取得できません。");
		return E_FAIL;
	}

	m_pCasManager = static_cast<TVCAS::ICasManager*>(pCreateInstance(__uuidof(TVCAS::ICasManager)));
	if (m_pCasManager == nullptr) {
		UnloadModule();
		OnError(E_FAIL, L"CASマネージャのインスタンスを作成できません。");
		return E_FAIL;
	}

	return S_OK;
}


STDMETHODIMP CFilterModuleImpl::UnloadModule()
{
	if (m_pCasManager != nullptr) {
		m_pCasManager->Release();
		m_pCasManager = nullptr;
	}

	if (m_hModule != nullptr) {
		::FreeLibrary(m_hModule);
		m_hModule = nullptr;
	}

	return S_OK;
}


STDMETHODIMP CFilterModuleImpl::IsModuleLoaded()
{
	return m_hModule != nullptr ? S_OK : S_FALSE;
}


STDMETHODIMP CFilterModuleImpl::GetModuleInfo(TVTest::Interface::FilterModuleInfo *pInfo)
{
	if (pInfo == nullptr)
		return E_POINTER;

	if (m_pCasManager == nullptr) {
		pInfo->Name = nullptr;
		pInfo->Version = nullptr;
		return E_UNEXPECTED;
	}

	BSTR Name, Version;

	Name = ::SysAllocString(m_CasModuleInfo.Name);
	Version = ::SysAllocString(m_CasModuleInfo.Version);

	if (Name == nullptr || Version == nullptr) {
		if (Name != nullptr)
			::SysFreeString(Name);
		if (Version != nullptr)
			::SysFreeString(Version);
		return E_OUTOFMEMORY;
	}

	pInfo->Name = Name;
	pInfo->Version = Version;

	return S_OK;
}


STDMETHODIMP CFilterModuleImpl::GetDeviceCount(int *pCount)
{
	if (pCount == nullptr)
		return E_POINTER;

	if (m_pCasManager == nullptr) {
		*pCount = 0;
		return E_UNEXPECTED;
	}

	*pCount = m_pCasManager->GetCasDeviceCount();

	return S_OK;
}


STDMETHODIMP CFilterModuleImpl::GetDevice(int Device, TVTest::Interface::IFilterDevice **ppDevice)
{
	if (ppDevice == nullptr)
		return E_POINTER;

	*ppDevice = nullptr;

	if (m_pCasManager == nullptr)
		return E_UNEXPECTED;

	TVCAS::ICasDevice *pDevice = m_pCasManager->OpenCasDevice(Device);
	if (pDevice == nullptr)
		return E_FAIL;

	*ppDevice = new CFilterDevice(pDevice);

	return S_OK;
}


STDMETHODIMP CFilterModuleImpl::IsDeviceAvailable(int Device, BOOL *pfAvailable)
{
	if (pfAvailable == nullptr)
		return E_POINTER;

	if (m_pCasManager == nullptr) {
		*pfAvailable = FALSE;
		return E_UNEXPECTED;
	}

	*pfAvailable = m_pCasManager->IsCasDeviceAvailable(Device);

	return S_OK;
}


STDMETHODIMP CFilterModuleImpl::CheckDeviceAvailability(int Device, BOOL *pfAvailable, BSTR *pMessage)
{
	if (pMessage != nullptr)
		*pMessage = nullptr;

	if (pfAvailable == nullptr)
		return E_POINTER;

	*pfAvailable = FALSE;

	if (m_pCasManager == nullptr)
		return E_UNEXPECTED;

	bool fAvailable;
	WCHAR szMessage[256];

	if (!m_pCasManager->CheckCasDeviceAvailability(Device, &fAvailable, szMessage, _countof(szMessage)))
		return E_FAIL;

	*pfAvailable = fAvailable;
	if (pMessage != nullptr && szMessage[0] != L'\0')
		*pMessage = ::SysAllocString(szMessage);

	return S_OK;
}


STDMETHODIMP CFilterModuleImpl::GetDefaultDevice(int *pDevice)
{
	if (pDevice == nullptr)
		return E_POINTER;

	if (m_pCasManager == nullptr) {
		*pDevice = -1;
		return E_UNEXPECTED;
	}

	*pDevice = m_pCasManager->GetDefaultCasDevice();

	return S_OK;
}


STDMETHODIMP CFilterModuleImpl::GetDeviceByID(TVTest::Interface::DeviceIDType DeviceID, int *pDevice)
{
	if (pDevice == nullptr)
		return E_POINTER;

	if (m_pCasManager == nullptr) {
		*pDevice = -1;
		return E_UNEXPECTED;
	}

	int Device = m_pCasManager->GetCasDeviceByID(DeviceID);
	*pDevice = Device;

	return Device >= 0 ? S_OK : E_INVALIDARG;
}


STDMETHODIMP CFilterModuleImpl::GetDeviceByName(LPCWSTR pszName, int *pDevice)
{
	if (pDevice == nullptr)
		return E_POINTER;

	if (m_pCasManager == nullptr) {
		*pDevice = -1;
		return E_UNEXPECTED;
	}

	int Device = m_pCasManager->GetCasDeviceByName(pszName);
	*pDevice = Device;

	return Device >= 0 ? S_OK : E_INVALIDARG;
}


STDMETHODIMP CFilterModuleImpl::OpenFilter(int Device, LPCWSTR pszName, TVTest::Interface::IFilter **ppFilter)
{
	return E_NOTIMPL;
}




class CFilterModule
	: public CFilterModuleImpl
	, protected CIUnknownImpl
{
public:
// IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject) override;
	STDMETHODIMP_(ULONG) AddRef() override { return AddRefImpl(); }
	STDMETHODIMP_(ULONG) Release() override { return ReleaseImpl(); }
};


STDMETHODIMP CFilterModule::QueryInterface(REFIID riid, void **ppvObject)
{
	if (ppvObject == nullptr)
		return E_POINTER;

	if (riid == IID_IUnknown || riid == __uuidof(IFilterModule)) {
		*ppvObject = static_cast<IFilterModule*>(this);
	} else {
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	AddRef();

	return S_OK;
}




class __declspec(uuid("781F8C21-18C6-4867-AD7D-73C954199CBA")) CPropertyPage
	: public IPropertyPage
	, protected CIUnknownImpl
{
public:
	CPropertyPage();

// IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject) override;
	STDMETHODIMP_(ULONG) AddRef() override { return AddRefImpl(); }
	STDMETHODIMP_(ULONG) Release() override { return ReleaseImpl(); }

// IPropertyPage
	STDMETHODIMP SetPageSite(IPropertyPageSite *pPageSite) override;
	STDMETHODIMP Activate(HWND hWndParent, LPCRECT pRect, BOOL bModal) override;
	STDMETHODIMP Deactivate() override;
	STDMETHODIMP GetPageInfo(PROPPAGEINFO *pPageInfo) override;
	STDMETHODIMP SetObjects(ULONG cObjects, IUnknown **ppUnk) override;
	STDMETHODIMP Show(UINT nCmdShow) override;
	STDMETHODIMP Move(LPCRECT pRect) override;
	STDMETHODIMP IsPageDirty() override;
	STDMETHODIMP Apply() override;
	STDMETHODIMP Help(LPCOLESTR pszHelpDir) override { return E_NOTIMPL; }
	STDMETHODIMP TranslateAccelerator(MSG *pMsg) override;

private:
	static const LPCTSTR DLG_PROP_NAME;
	static const UINT WM_APP_UPDATECONTROLS = WM_APP;

	IPropertyPageSite *m_pPageSite;
	ICasProcessor *m_pCasProcessor;
	HWND m_hwnd;
	HWND m_hwndParking;
	bool m_fDirty;

	~CPropertyPage();
	bool LoadPage(HWND hwndParent);
	HWND GetParkingWindow();
	void MakeDirty();
	void UpdateControls();
	void BenchmarkTest();
	static INT_PTR CALLBACK DlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
	static CPropertyPage *GetThis(HWND hDlg) {
		return static_cast<CPropertyPage*>(::GetProp(hDlg, DLG_PROP_NAME));
	}
};


const LPCTSTR CPropertyPage::DLG_PROP_NAME = TEXT("0A32318D-6C2D-4fbf-9065-02B84719FBB4");


CPropertyPage::CPropertyPage()
	: m_pPageSite(nullptr)
	, m_pCasProcessor(nullptr)
	, m_hwnd(nullptr)
	, m_hwndParking(nullptr)
	, m_fDirty(false)
{
}


CPropertyPage::~CPropertyPage()
{
	if (m_hwnd != nullptr)
		::DestroyWindow(m_hwnd);
	if (m_hwndParking != nullptr)
		::DestroyWindow(m_hwndParking);
	if (m_pPageSite != nullptr)
		m_pPageSite->Release();
	if (m_pCasProcessor != nullptr)
		m_pCasProcessor->Release();
}


STDMETHODIMP CPropertyPage::QueryInterface(REFIID riid, void **ppvObject)
{
	if (ppvObject == nullptr)
		return E_POINTER;

	if (riid == IID_IUnknown || riid == __uuidof(IPropertyPage)) {
		*ppvObject = static_cast<IPropertyPage*>(this);
	} else {
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	AddRef();

	return S_OK;
}


STDMETHODIMP CPropertyPage::SetPageSite(IPropertyPageSite *pPageSite)
{
	if (m_pPageSite != nullptr)
		m_pPageSite->Release();
	m_pPageSite = pPageSite;
	if (m_pPageSite != nullptr)
		m_pPageSite->AddRef();

	return S_OK;
}


STDMETHODIMP CPropertyPage::Activate(HWND hWndParent, LPCRECT pRect, BOOL bModal)
{
	if (pRect == nullptr)
		return E_POINTER;

	if (!LoadPage(hWndParent))
		return E_FAIL;

	Move(pRect);

	return S_OK;
}


STDMETHODIMP CPropertyPage::Deactivate()
{
	if (m_hwnd == nullptr)
		return E_UNEXPECTED;

	::DestroyWindow(m_hwnd);
	m_hwnd = nullptr;

	return S_OK;
}


STDMETHODIMP CPropertyPage::GetPageInfo(PROPPAGEINFO *pPageInfo)
{
	if (pPageInfo == nullptr)
		return E_POINTER;

	if (!LoadPage(GetParkingWindow()))
		return E_FAIL;

	RECT rc;
	::GetWindowRect(m_hwnd, &rc);

	pPageInfo->cb = sizeof(PROPPAGEINFO);
	pPageInfo->pszTitle = AllocateOleString(OLESTR("CasProcessor"));
	pPageInfo->size.cx = rc.right - rc.left;
	pPageInfo->size.cy = rc.bottom - rc.top;
	pPageInfo->pszDocString = nullptr;
	pPageInfo->pszHelpFile = nullptr;
	pPageInfo->dwHelpContext = 0;

	return S_OK;
}


STDMETHODIMP CPropertyPage::SetObjects(ULONG cObjects, IUnknown **ppUnk)
{
	if (cObjects > 0 && ppUnk == nullptr)
		return E_POINTER;

	if (m_pCasProcessor != nullptr) {
		m_pCasProcessor->Release();
		m_pCasProcessor = nullptr;
	}

	for (ULONG i = 0; i < cObjects; i++) {
		ICasProcessor *pCasProcessor;
		if (SUCCEEDED(ppUnk[i]->QueryInterface(IID_PPV_ARGS(&pCasProcessor)))) {
			m_pCasProcessor = pCasProcessor;
			break;
		}
	}

	if (m_hwnd != nullptr)
		::SendMessage(m_hwnd, WM_APP_UPDATECONTROLS, 0, 0);

	return S_OK;
}


STDMETHODIMP CPropertyPage::Show(UINT nCmdShow)
{
	if (m_hwnd == nullptr)
		return E_UNEXPECTED;

	::ShowWindow(m_hwnd, nCmdShow);

	return S_OK;
}


STDMETHODIMP CPropertyPage::Move(LPCRECT pRect)
{
	if (pRect == nullptr)
		return E_POINTER;
	if (m_hwnd == nullptr)
		return E_UNEXPECTED;

	::MoveWindow(m_hwnd, pRect->left, pRect->top,
				 pRect->right - pRect->left, pRect->bottom - pRect->top, TRUE);

	return S_OK;
}


STDMETHODIMP CPropertyPage::IsPageDirty()
{
	return m_fDirty ? S_OK : S_FALSE;
}


STDMETHODIMP CPropertyPage::Apply()
{
	if (m_hwnd == nullptr)
		return E_UNEXPECTED;

	if (m_fDirty) {
		if (m_pCasProcessor != nullptr) {
			m_pCasProcessor->SetSpecificServiceDecoding(
				::IsDlgButtonChecked(m_hwnd, IDC_PROPERTIES_SPECIFICSERVICEDECODING) == BST_CHECKED);
			m_pCasProcessor->SetEnableContract(
				::IsDlgButtonChecked(m_hwnd, IDC_PROPERTIES_ENABLEEMMPROCESS) == BST_CHECKED);
			int Sel = (int)::SendDlgItemMessage(m_hwnd, IDC_PROPERTIES_INSTRUCTION, CB_GETCURSEL, 0, 0);
			if (Sel >= 0) {
				m_pCasProcessor->SetInstruction(
					(int)::SendDlgItemMessage(m_hwnd, IDC_PROPERTIES_INSTRUCTION, CB_GETITEMDATA, Sel, 0));
			}
		}

		m_fDirty = false;
		if (m_pPageSite != nullptr)
			m_pPageSite->OnStatusChange(PROPPAGESTATUS_DIRTY);
	}

	return S_OK;
}


STDMETHODIMP CPropertyPage::TranslateAccelerator(MSG *pMsg)
{
	if (pMsg == nullptr)
		return E_POINTER;
	if (m_hwnd == nullptr)
		return E_UNEXPECTED;

	return ::IsDialogMessage(m_hwnd, pMsg) ? S_OK : S_FALSE;
}


bool CPropertyPage::LoadPage(HWND hwndParent)
{
	if (m_hwnd == nullptr) {
		m_hwnd = ::CreateDialogParam(
			g_hinstDLL, MAKEINTRESOURCE(IDD_PROPERTIES),
			hwndParent, DlgProc, reinterpret_cast<LPARAM>(this));
		if (m_hwnd == nullptr)
			return false;
	} else {
		if (::GetParent(m_hwnd) != hwndParent)
			::SetParent(m_hwnd, hwndParent);
	}

	return true;
}


HWND CPropertyPage::GetParkingWindow()
{
	if (m_hwndParking == nullptr) {
		m_hwndParking = ::CreateWindowEx(
			0, TEXT("STATIC"), TEXT(""), WS_POPUP,
			0, 0, 0, 0, nullptr, nullptr, g_hinstDLL, nullptr);
	}

	return m_hwndParking;
}


void CPropertyPage::MakeDirty()
{
	m_fDirty = true;
	if (m_pPageSite != nullptr)
		m_pPageSite->OnStatusChange(PROPPAGESTATUS_DIRTY | PROPPAGESTATUS_VALIDATE);
}


void CPropertyPage::UpdateControls()
{
	if (m_hwnd == nullptr || m_pCasProcessor == nullptr)
		return;

	BOOL fEnable;
	fEnable = FALSE;
	m_pCasProcessor->GetSpecificServiceDecoding(&fEnable);
	::CheckDlgButton(m_hwnd, IDC_PROPERTIES_SPECIFICSERVICEDECODING,
					 fEnable ? BST_CHECKED : BST_UNCHECKED);
	fEnable = FALSE;
	m_pCasProcessor->GetEnableContract(&fEnable);
	::CheckDlgButton(m_hwnd, IDC_PROPERTIES_ENABLEEMMPROCESS,
					 fEnable ? BST_CHECKED : BST_UNCHECKED);

	bool fEnableInstructionSettings = false;
	UINT AvailableInstructions;
	if (SUCCEEDED(m_pCasProcessor->GetAvailableInstructions(&AvailableInstructions))) {
		int Instruction;
		m_pCasProcessor->GetInstruction(&Instruction);
		int Sel = -1;
		for (int i = 0; AvailableInstructions >> i != 0; i++) {
			if (((AvailableInstructions >> i) & 1) != 0) {
				BSTR Name;
				if (SUCCEEDED(m_pCasProcessor->GetInstructionName(i, &Name))) {
					int Index=(int)::SendDlgItemMessage(m_hwnd, IDC_PROPERTIES_INSTRUCTION,
														CB_ADDSTRING, 0, (LPARAM)Name);
					::SysFreeString(Name);
					::SendDlgItemMessage(m_hwnd, IDC_PROPERTIES_INSTRUCTION,
										 CB_SETITEMDATA, Index, i);
					if (i == Instruction)
						Sel = Index;
				}
			}
		}
		::SendDlgItemMessage(m_hwnd, IDC_PROPERTIES_INSTRUCTION, CB_SETCURSEL, Sel, 0);
		fEnableInstructionSettings = true;
	}
	::EnableWindow(::GetDlgItem(m_hwnd, IDC_PROPERTIES_INSTRUCTION_LABEL),
				   fEnableInstructionSettings);
	::EnableWindow(::GetDlgItem(m_hwnd, IDC_PROPERTIES_INSTRUCTION),
				   fEnableInstructionSettings);
	::EnableWindow(::GetDlgItem(m_hwnd, IDC_PROPERTIES_BENCHMARKTEST),
				   fEnableInstructionSettings);
	::SetDlgItemText(m_hwnd, IDC_PROPERTIES_INSTRUCTION_NOTE,
		fEnableInstructionSettings ?
			TEXT("※この設定は次回反映されます。") :
			TEXT("※この設定はモジュールが読み込まれている時のみ行えます。"));

	BSTR bstr;
	if (SUCCEEDED(m_pCasProcessor->GetCardReaderName(&bstr))) {
		::SetDlgItemTextW(m_hwnd, IDC_PROPERTIES_CARDREADERNAME, bstr);
		::SysFreeString(bstr);
	}
	if (SUCCEEDED(m_pCasProcessor->GetCardID(&bstr))) {
		::SetDlgItemTextW(m_hwnd, IDC_PROPERTIES_CARDID, bstr);
		::SysFreeString(bstr);
	}
	if (SUCCEEDED(m_pCasProcessor->GetCardVersion(&bstr))) {
		::SetDlgItemTextW(m_hwnd, IDC_PROPERTIES_CARDVERSION, bstr);
		::SysFreeString(bstr);
	}
}


void CPropertyPage::BenchmarkTest()
{
	if (m_pCasProcessor == nullptr)
		return;

	UINT AvailableInstructions;
	if (FAILED(m_pCasProcessor->GetAvailableInstructions(&AvailableInstructions))
		|| AvailableInstructions == 0)
		return;

	HCURSOR hcurOld=::SetCursor(LoadCursor(NULL,IDC_WAIT));

	static const DWORD BENCHMARK_ROUND = 200000;
	DWORD BenchmarkCount = 0, MaxTime = 0;
	DWORD Times[32];
	::ZeroMemory(Times,sizeof(Times));

	do {
		for (int i = 0; AvailableInstructions >> i != 0; i++) {
			if (((AvailableInstructions >> i) & 1) != 0) {
				DWORD Time;
				if (SUCCEEDED(m_pCasProcessor->BenchmarkTest(i, BENCHMARK_ROUND, &Time))) {
					Times[i] += Time;
					if (Times[i] > MaxTime)
						MaxTime = Times[i];
				}
			}
		}
		BenchmarkCount += BENCHMARK_ROUND;
	} while (MaxTime < 1500);

	::SetCursor(hcurOld);

	TCHAR szText[1024];
	int Pos = 0;

	TCHAR szCPU[256];
	DWORD Type = REG_SZ, Size = sizeof(szCPU);
	if (::SHGetValue(HKEY_LOCAL_MACHINE,
					 TEXT("HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\0"),
					 TEXT("ProcessorNameString"),
					 &Type, szCPU, &Size) == ERROR_SUCCESS
		&& Type == REG_SZ)
		StringFormatAppend(szText, _countof(szText), &Pos, TEXT("%s\n"), szCPU);

	StringFormatAppend(szText, _countof(szText), &Pos,
					   TEXT("%u 回の実行に掛かった時間\n\n"), BenchmarkCount);

	DWORD NormalTime = Times[0];
	DWORD MinTime = 0xFFFFFFFF;
	int Fastest=0;

	for (int i = 0; AvailableInstructions >> i != 0; i++) {
		if (((AvailableInstructions >> i) & 1) != 0) {
			const DWORD Time = Times[i];
			BSTR Name;

			m_pCasProcessor->GetInstructionName(i, &Name);
			StringFormatAppend(szText, _countof(szText), &Pos,
							   TEXT("%s : %u ms (%d パケット/秒)"),
				Name, Time, ::MulDiv(BenchmarkCount, 1000, Time));
			::SysFreeString(Name);
			if (i > 0 && NormalTime > 0 && Time > 0) {
				int Percentage;
				if (NormalTime >= Time)
					Percentage = (int)(NormalTime * 100 / Time) - 100;
				else
					Percentage = -(int)((Time * 100 / NormalTime) - 100);
				StringFormatAppend(szText, _countof(szText), &Pos,
								   TEXT(" (高速化される割合 %d %%)"), Percentage);
			}
			StringFormatAppend(szText, _countof(szText), &Pos, TEXT("\n"));

			if (Time < MinTime) {
				MinTime = Time;
				Fastest = i;
			}
		}
	}

	BSTR Name;
	m_pCasProcessor->GetInstructionName(Fastest, &Name);
	StringFormatAppend(szText, _countof(szText), &Pos,
					   TEXT("\n%s にすることをお勧めします。"), Name);
	::SysFreeString(Name);

	::MessageBox(m_hwnd,
				 szText, TEXT("ベンチマークテスト結果"),
				 MB_OK | MB_ICONINFORMATION);
}


INT_PTR CALLBACK CPropertyPage::DlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg) {
	case WM_INITDIALOG:
		{
			CPropertyPage *pThis = reinterpret_cast<CPropertyPage*>(lParam);

			pThis->m_hwnd = hDlg;
			::SetProp(hDlg, DLG_PROP_NAME, pThis);
			::SendMessage(hDlg, WM_APP_UPDATECONTROLS, 0, 0);
		}
		return TRUE;

	case WM_COMMAND:
		switch (LOWORD(wParam)) {
		case IDC_PROPERTIES_SPECIFICSERVICEDECODING:
		case IDC_PROPERTIES_ENABLEEMMPROCESS:
			{
				CPropertyPage *pThis = GetThis(hDlg);

				if (pThis != nullptr)
					pThis->MakeDirty();
			}
			return TRUE;

		case IDC_PROPERTIES_INSTRUCTION:
			if (HIWORD(wParam) == CBN_SELCHANGE) {
				CPropertyPage *pThis = GetThis(hDlg);

				if (pThis != nullptr)
					pThis->MakeDirty();
			}
			return TRUE;

		case IDC_PROPERTIES_BENCHMARKTEST:
			if (::MessageBox(hDlg,
					TEXT("ベンチマークテストを開始します。\n")
					TEXT("終了するまで操作は行わないようにしてください。\n")
					TEXT("結果はばらつきがありますので、数回実行してください。"),
					TEXT("ベンチマークテスト"),
					MB_OKCANCEL | MB_ICONINFORMATION) == IDOK) {
				CPropertyPage *pThis = GetThis(hDlg);

				if (pThis != nullptr)
					pThis->BenchmarkTest();
			}
			return TRUE;
		}
		return TRUE;

	case WM_APP_UPDATECONTROLS:
		{
			CPropertyPage *pThis = GetThis(hDlg);

			::SendDlgItemMessage(hDlg, IDC_PROPERTIES_INSTRUCTION, CB_RESETCONTENT, 0, 0);
			::SetDlgItemText(hDlg, IDC_PROPERTIES_CARDREADERNAME, TEXT(""));
			::SetDlgItemText(hDlg, IDC_PROPERTIES_CARDID, TEXT(""));
			::SetDlgItemText(hDlg, IDC_PROPERTIES_CARDVERSION, TEXT(""));

			if (pThis != nullptr)
				pThis->UpdateControls();
		}
		return TRUE;

	case WM_DESTROY:
		{
			CPropertyPage *pThis = GetThis(hDlg);

			if (pThis != nullptr) {
				pThis->m_hwnd = nullptr;
				::RemoveProp(hDlg, DLG_PROP_NAME);
			}
		}
		return TRUE;
	}

	return FALSE;
}




class CCasProcessor
	: public CFilterModuleImpl
	, public TVTest::Interface::ITSProcessor
	, public TVTest::Interface::IFilterManager
	, public IPersistPropertyBag
	, public ISpecifyPropertyPages2
	, public ICasProcessor
	, protected CIUnknownImpl
{
public:
	CCasProcessor();

// IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject) override;
	STDMETHODIMP_(ULONG) AddRef() override { return AddRefImpl(); }
	STDMETHODIMP_(ULONG) Release() override { return ReleaseImpl(); }

// ITSProcessor
	STDMETHODIMP GetGuid(GUID *pGuid) override;
	STDMETHODIMP GetName(BSTR *pName) override;
	STDMETHODIMP Initialize(TVTest::Interface::IStreamingClient *pClient) override;
	STDMETHODIMP Finalize() override;
	STDMETHODIMP StartStreaming(TVTest::Interface::ITSOutput *pOutput) override;
	STDMETHODIMP StopStreaming() override;
	STDMETHODIMP InputPacket(TVTest::Interface::ITSPacket *pPacket) override;
	STDMETHODIMP Reset() override;
	STDMETHODIMP SetEnableProcessing(BOOL fEnable) override;
	STDMETHODIMP GetEnableProcessing(BOOL *pfEnable) override;
	STDMETHODIMP SetActiveServiceID(WORD ServiceID) override;
	STDMETHODIMP GetActiveServiceID(WORD *pServiceID) override;

// IFilterModule
	STDMETHODIMP LoadModule(LPCWSTR pszName) override;
	STDMETHODIMP UnloadModule() override;
	STDMETHODIMP OpenFilter(int Device, LPCWSTR pszName, TVTest::Interface::IFilter **ppFilter) override;

// IFilterManager
	STDMETHODIMP EnumModules(TVTest::Interface::IEnumFilterModule **ppEnumModule) override;
	STDMETHODIMP CreateModule(TVTest::Interface::IFilterModule **ppModule) override;

// IPersistPropertyBag
	STDMETHODIMP GetClassID(CLSID *pClassID) override;
	STDMETHODIMP InitNew() override;
	STDMETHODIMP Load(IPropertyBag *pPropBag, IErrorLog *pErrorLog) override;
	STDMETHODIMP Save(IPropertyBag *pPropBag, BOOL fClearDirty, BOOL fSaveAllProperties) override;

// ISpecifyPropertyPages
	STDMETHODIMP GetPages(CAUUID *pPages) override;

// ISpecifyPropertyPages2
	STDMETHODIMP CreatePage(const GUID &guid, IPropertyPage **ppPage) override;

// ICasProcessor
	STDMETHODIMP SetSpecificServiceDecoding(BOOL fSpecificService) override;
	STDMETHODIMP GetSpecificServiceDecoding(BOOL *pfSpecificService) override;
	STDMETHODIMP SetEnableContract(BOOL fEnable) override;
	STDMETHODIMP GetEnableContract(BOOL *pfEnable) override;
	STDMETHODIMP GetCardReaderName(BSTR *pName) override;
	STDMETHODIMP GetCardID(BSTR *pID) override;
	STDMETHODIMP GetCardVersion(BSTR *pVersion) override;
	STDMETHODIMP GetInstructionName(int Instruction, BSTR *pName) override;
	STDMETHODIMP GetAvailableInstructions(UINT *pAvailableInstructions) override;
	STDMETHODIMP SetInstruction(int Instruction) override;
	STDMETHODIMP GetInstruction(int *pInstruction) override;
	STDMETHODIMP BenchmarkTest(int Instruction, DWORD Round, DWORD *pTime) override;

private:
	class CCasClient
		: public TVCAS::ICasClient
		, protected TVCAS::Helper::CBaseImpl
	{
	public:
		CCasClient(CCasProcessor *pCasProcessor);

	// TVCAS::IBase
		TVCAS_DECLARE_BASE
		LPCWSTR GetName() const override { return L"CasProcessor"; }

	// TVCAS::ICasClient
		LRESULT OnEvent(UINT Event, void *pParam) override;
		LRESULT OnError(const TVCAS::ErrorInfo *pInfo) override;
		void OutLog(TVCAS::LogType Type, LPCWSTR pszMessage) override;

	protected:
		~CCasClient();

		CCasProcessor *m_pCasProcessor;
	};

	static const GUID m_ProcessorGuid;
	static const bool m_fSpecificServiceDecodingDefault = true;
	static const bool m_fEnableContractDefault = true;

	TVTest::Interface::IStreamingClient *m_pClient;
	TVTest::Interface::ITSOutput *m_pOutput;
	CLock m_Lock;
	bool m_fFilterOpened;
	WORD m_ActiveServiceID;
	bool m_fEnableProcessing;
	bool m_fDirty;
	bool m_fSpecificServiceDecoding;
	bool m_fEnableContract;
	int m_Instruction;

	~CCasProcessor();
	HRESULT OnError(HRESULT hr, LPCWSTR pszText, LPCWSTR pszAdvise = nullptr, LPCWSTR pszSystemMessage = nullptr) override;
	void OutLog(TVTest::Interface::LogType Type, LPCWSTR pszFormat, ...);
	void OnCasEvent(UINT Event, void *pParam);
};


// {24690A0B-E97A-422e-BE98-09F539F87EE4}
const GUID CCasProcessor::m_ProcessorGuid = {
	0x24690a0b, 0xe97a, 0x422e, { 0xbe, 0x98, 0x9, 0xf5, 0x39, 0xf8, 0x7e, 0xe4 }
};


CCasProcessor::CCasProcessor()
	: m_pClient(nullptr)
	, m_pOutput(nullptr)
	, m_fFilterOpened(false)
	, m_ActiveServiceID(0)
	, m_fEnableProcessing(true)
	, m_fDirty(false)
	, m_fSpecificServiceDecoding(m_fSpecificServiceDecodingDefault)
	, m_fEnableContract(m_fEnableContractDefault)
	, m_Instruction(-1)
{
}


CCasProcessor::~CCasProcessor()
{
	Finalize();
	UnloadModule();
}


STDMETHODIMP CCasProcessor::QueryInterface(REFIID riid, void **ppvObject)
{
	if (ppvObject == nullptr)
		return E_POINTER;

	if (riid == IID_IUnknown || riid == __uuidof(ITSProcessor)) {
		*ppvObject = static_cast<ITSProcessor*>(this);
	} else if (riid == __uuidof(IFilterManager)) {
		*ppvObject = static_cast<IFilterManager*>(this);
	} else if (riid == __uuidof(IFilterModule)) {
		*ppvObject = static_cast<IFilterModule*>(this);
	} else if (riid == __uuidof(IPersistPropertyBag)) {
		*ppvObject = static_cast<IPersistPropertyBag*>(this);
	} else if (riid == __uuidof(ISpecifyPropertyPages)) {
		*ppvObject = static_cast<ISpecifyPropertyPages*>(this);
	} else if (riid == __uuidof(ISpecifyPropertyPages2)) {
		*ppvObject = static_cast<ISpecifyPropertyPages2*>(this);
	} else if (riid == __uuidof(ICasProcessor)) {
		*ppvObject = static_cast<ICasProcessor*>(this);
	} else {
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}

	AddRef();

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetGuid(GUID *pGuid)
{
	if (pGuid == nullptr)
		return E_POINTER;

	*pGuid = m_ProcessorGuid;

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetName(BSTR *pName)
{
	if (pName == nullptr)
		return E_POINTER;

	*pName = ::SysAllocString(L"CasProcessor");

	if (*pName == nullptr)
		return E_OUTOFMEMORY;

	return S_OK;
}


STDMETHODIMP CCasProcessor::Initialize(TVTest::Interface::IStreamingClient *pClient)
{
	if (pClient == nullptr)
		return E_POINTER;

	if (m_pClient != nullptr)
		return E_UNEXPECTED;

	m_pClient = pClient;
	m_pClient->AddRef();

	return S_OK;
}


STDMETHODIMP CCasProcessor::Finalize()
{
	StopStreaming();

	if (m_pClient != nullptr) {
		m_pClient->Release();
		m_pClient = nullptr;
	}

	return S_OK;
}


STDMETHODIMP CCasProcessor::StartStreaming(TVTest::Interface::ITSOutput *pOutput)
{
	m_pOutput = pOutput;

	if (m_pOutput != nullptr)
		m_pOutput->AddRef();

	return S_OK;
}


STDMETHODIMP CCasProcessor::StopStreaming()
{
	if (m_pOutput != nullptr) {
		m_pOutput->Release();
		m_pOutput = nullptr;
	}

	return S_OK;
}


STDMETHODIMP CCasProcessor::InputPacket(TVTest::Interface::ITSPacket *pPacket)
{
	if (pPacket == nullptr)
		return E_POINTER;

	{
		CBlockLock Lock(m_Lock);

		if (m_pCasManager != nullptr && m_fFilterOpened) {
			BYTE *pData;
			HRESULT hr = pPacket->GetData(&pData);
			if (SUCCEEDED(hr)) {
				ULONG Size;
				hr = pPacket->GetSize(&Size);
				if (SUCCEEDED(hr)) {
					BYTE OldScramblingCtrl = (pData[3] & 0xC0);
					m_pCasManager->ProcessPacket(pData, Size);
					if ((pData[3] & 0xC0) != OldScramblingCtrl)
						pPacket->SetModified(TRUE);
				}
			}
		}
	}

	if (m_pOutput != NULL)
		m_pOutput->OutputPacket(pPacket);

	return S_OK;
}


STDMETHODIMP CCasProcessor::Reset()
{
	CBlockLock Lock(m_Lock);

	if (m_pCasManager != nullptr)
		m_pCasManager->Reset();

	return S_OK;
}


STDMETHODIMP CCasProcessor::SetEnableProcessing(BOOL fEnable)
{
	m_fEnableProcessing = fEnable != FALSE;

	if (m_pCasManager != nullptr)
		m_pCasManager->EnableDescramble(m_fEnableProcessing);

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetEnableProcessing(BOOL *pfEnable)
{
	if (pfEnable == nullptr)
		return E_POINTER;

	*pfEnable = m_fEnableProcessing;

	return S_OK;
}


STDMETHODIMP CCasProcessor::SetActiveServiceID(WORD ServiceID)
{
	if (m_pCasManager != nullptr && m_fSpecificServiceDecoding)
		m_pCasManager->SetDescrambleServiceID(ServiceID);

	m_ActiveServiceID = ServiceID;

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetActiveServiceID(WORD *pServiceID)
{
	if (pServiceID == nullptr)
		return E_POINTER;

	*pServiceID = m_ActiveServiceID;

	return S_OK;
}


STDMETHODIMP CCasProcessor::LoadModule(LPCWSTR pszName)
{
	CBlockLock Lock(m_Lock);

	HRESULT hr = CFilterModuleImpl::LoadModule(pszName);
	if (FAILED(hr))
		return hr;

	CCasClient *pCasClient = new CCasClient(this);
	bool fResult = m_pCasManager->Initialize(pCasClient);
	pCasClient->Release();
	if (!fResult) {
		UnloadModule();
		OnError(E_FAIL, L"CASマネージャの初期化ができません。");
		return E_FAIL;
	}

	m_pCasManager->EnableDescramble(m_fEnableProcessing);
	m_pCasManager->EnableContract(m_fEnableContract);
	if (m_Instruction >= 0) {
		if (((1U << m_Instruction) & m_pCasManager->GetAvailableInstructions()) != 0)
			m_pCasManager->SetInstruction(m_Instruction);
	} else {
		m_Instruction = m_pCasManager->GetInstruction();
	}

	return S_OK;
}


STDMETHODIMP CCasProcessor::UnloadModule()
{
	CBlockLock Lock(m_Lock);

	CFilterModuleImpl::UnloadModule();

	m_fFilterOpened = false;

	return S_OK;
}


STDMETHODIMP CCasProcessor::OpenFilter(int Device, LPCWSTR pszName, TVTest::Interface::IFilter **ppFilter)
{
	if (ppFilter == nullptr)
		return E_POINTER;

	*ppFilter = nullptr;

	CBlockLock Lock(m_Lock);

	if (m_pCasManager == nullptr)
		return E_UNEXPECTED;

	if (!m_pCasManager->OpenCasCard(Device, (pszName != nullptr && pszName[0] != L'\0') ? pszName : nullptr))
		return E_FAIL;

	WCHAR szName[MAX_PATH];
	if (m_pCasManager->GetCasCardName(szName, _countof(szName)) > 0) {
		OutLog(TVTest::Interface::LOG_INFO, L"カードリーダー \"%s\" をオープンしました。", szName);
		TVCAS::CasCardInfo CardInfo;
		if (m_pCasManager->GetCasCardInfo(&CardInfo)) {
			OutLog(TVTest::Interface::LOG_INFO, L"(カードID %s / カード識別 %c%03d)",
				   CardInfo.CardIDText, CardInfo.CardManufacturerID, CardInfo.CardVersion);
		}
	}

	*ppFilter = new CFilter(m_pCasManager, Device);

	m_fFilterOpened = true;

	return S_OK;
}


STDMETHODIMP CCasProcessor::EnumModules(TVTest::Interface::IEnumFilterModule **ppEnumModule)
{
	if (ppEnumModule == nullptr)
		return E_POINTER;

	*ppEnumModule = new CEnumFilterModule;

	return S_OK;
}


STDMETHODIMP CCasProcessor::CreateModule(TVTest::Interface::IFilterModule **ppModule)
{
	if (ppModule == nullptr)
		return E_POINTER;

	*ppModule = new CFilterModule;

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetClassID(CLSID *pClassID)
{
	return GetGuid(pClassID);
}


STDMETHODIMP CCasProcessor::InitNew()
{
	return S_OK;
}


STDMETHODIMP CCasProcessor::Load(IPropertyBag *pPropBag, IErrorLog *pErrorLog)
{
	if (pPropBag == nullptr)
		return E_POINTER;

	VARIANT var;

	::VariantInit(&var);

	var.vt = VT_BOOL;
	if (SUCCEEDED(pPropBag->Read(L"EnableContract", &var, nullptr)))
		SetEnableContract(var.boolVal == VARIANT_TRUE);

	var.vt = VT_BOOL;
	if (SUCCEEDED(pPropBag->Read(L"SpecificServiceDecoding", &var, nullptr)))
		SetSpecificServiceDecoding(var.boolVal == VARIANT_TRUE);

	var.vt = VT_I4;
	if (SUCCEEDED(pPropBag->Read(L"Instruction", &var, nullptr)))
		SetInstruction(var.lVal);

	return S_OK;
}


STDMETHODIMP CCasProcessor::Save(IPropertyBag *pPropBag, BOOL fClearDirty, BOOL fSaveAllProperties)
{
	if (pPropBag == nullptr)
		return E_POINTER;

	VARIANT var;

	::VariantInit(&var);

	if (fSaveAllProperties || m_fEnableContract != m_fEnableContractDefault) {
		var.vt = VT_BOOL;
		var.boolVal = m_fEnableContract ? VARIANT_TRUE : VARIANT_FALSE;
		pPropBag->Write(L"EnableContract", &var);
	}

	if (fSaveAllProperties || m_fSpecificServiceDecoding != m_fSpecificServiceDecodingDefault) {
		var.vt = VT_BOOL;
		var.boolVal = m_fSpecificServiceDecoding ? VARIANT_TRUE : VARIANT_FALSE;
		pPropBag->Write(L"SpecificServiceDecoding", &var);
	}

	if (fSaveAllProperties || m_Instruction >= 0) {
		var.vt = VT_I4;
		var.lVal = m_Instruction;
		pPropBag->Write(L"Instruction", &var);
	}

	if (fClearDirty)
		m_fDirty = false;

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetPages(CAUUID *pPages)
{
	if (pPages == nullptr)
		return E_POINTER;

	pPages->pElems = static_cast<GUID*>(::CoTaskMemAlloc(sizeof(GUID)));
	if (pPages->pElems == nullptr) {
		pPages->cElems = 0;
		return E_OUTOFMEMORY;
	}

	pPages->cElems = 1;
	pPages->pElems[0] = __uuidof(CPropertyPage);

	return S_OK;
}


STDMETHODIMP CCasProcessor::CreatePage(const GUID &guid, IPropertyPage **ppPage)
{
	if (ppPage == nullptr)
		return E_POINTER;

	if (guid == __uuidof(CPropertyPage)) {
		*ppPage = new CPropertyPage;
	} else {
		return E_INVALIDARG;
	}

	return S_OK;
}


STDMETHODIMP CCasProcessor::SetSpecificServiceDecoding(BOOL fSpecificService)
{
	bool fSpecificServiceDecoding = fSpecificService != FALSE;

	if (m_fSpecificServiceDecoding != fSpecificServiceDecoding) {
		m_fSpecificServiceDecoding = fSpecificServiceDecoding;
		m_fDirty = true;

		if (m_pCasManager != nullptr) {
			m_pCasManager->SetDescrambleServiceID(
				m_fSpecificServiceDecoding ? m_ActiveServiceID : 0);
		}
	}

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetSpecificServiceDecoding(BOOL *pfSpecificService)
{
	if (pfSpecificService == nullptr)
		return E_POINTER;

	*pfSpecificService = m_fSpecificServiceDecoding;

	return S_OK;
}


STDMETHODIMP CCasProcessor::SetEnableContract(BOOL fEnable)
{
	bool fEnableContract = fEnable != FALSE;

	if (m_fEnableContract != fEnableContract) {
		m_fEnableContract = fEnableContract;
		m_fDirty = true;

		if (m_pCasManager != nullptr)
			m_pCasManager->EnableContract(m_fEnableContract);
	}

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetEnableContract(BOOL *pfEnable)
{
	if (pfEnable == nullptr)
		return E_POINTER;

	*pfEnable = m_fEnableContract;

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetCardReaderName(BSTR *pName)
{
	if (pName == nullptr)
		return E_POINTER;

	*pName = nullptr;

	if (m_pCasManager == nullptr)
		return E_UNEXPECTED;

	WCHAR szName[MAX_PATH];

	if (m_pCasManager->GetCasCardName(szName, _countof(szName)) < 1)
		return E_FAIL;

	*pName = ::SysAllocString(szName);
	if (*pName == nullptr)
		return E_OUTOFMEMORY;

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetCardID(BSTR *pID)
{
	if (pID == nullptr)
		return E_POINTER;

	*pID = nullptr;

	if (m_pCasManager == nullptr)
		return E_UNEXPECTED;

	TVCAS::CasCardInfo CardInfo;

	if (!m_pCasManager->GetCasCardInfo(&CardInfo))
		return E_FAIL;

	*pID = ::SysAllocString(CardInfo.CardIDText);
	if (*pID == nullptr)
		return E_OUTOFMEMORY;

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetCardVersion(BSTR *pVersion)
{
	if (pVersion == nullptr)
		return E_POINTER;

	*pVersion = nullptr;

	if (m_pCasManager == nullptr)
		return E_UNEXPECTED;

	TVCAS::CasCardInfo CardInfo;

	if (!m_pCasManager->GetCasCardInfo(&CardInfo))
		return E_FAIL;

	WCHAR szVersion[8];
	::wnsprintfW(szVersion, _countof(szVersion),
				 L"%c%03d",
				 CardInfo.CardManufacturerID,
				 CardInfo.CardVersion);

	*pVersion = ::SysAllocString(szVersion);
	if (*pVersion == nullptr)
		return E_OUTOFMEMORY;

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetInstructionName(int Instruction, BSTR *pName)
{
	if (pName == nullptr)
		return E_POINTER;

	*pName = nullptr;

	if (m_pCasManager == nullptr)
		return E_UNEXPECTED;

	WCHAR szName[256];

	if (m_pCasManager->GetInstructionName(Instruction, szName, _countof(szName)) < 1)
		return E_FAIL;

	*pName = ::SysAllocString(szName);
	if (*pName == nullptr)
		return E_OUTOFMEMORY;

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetAvailableInstructions(UINT *pAvailableInstructions)
{
	if (pAvailableInstructions == nullptr)
		return E_POINTER;

	if (m_pCasManager == nullptr) {
		*pAvailableInstructions = 0;
		return E_UNEXPECTED;
	}

	*pAvailableInstructions = m_pCasManager->GetAvailableInstructions();

	return S_OK;
}


STDMETHODIMP CCasProcessor::SetInstruction(int Instruction)
{
	if (m_Instruction != Instruction) {
		if (m_pCasManager != nullptr) {
			if (!m_pCasManager->SetInstruction(Instruction))
				return E_FAIL;
		}

		m_Instruction = Instruction;
		m_fDirty = true;
	}

	return S_OK;
}


STDMETHODIMP CCasProcessor::GetInstruction(int *pInstruction)
{
	if (pInstruction == nullptr)
		return E_POINTER;

	if (m_pCasManager == nullptr) {
		*pInstruction = -1;
		return E_UNEXPECTED;
	}

	*pInstruction = m_pCasManager->GetInstruction();

	return S_OK;
}


STDMETHODIMP CCasProcessor::BenchmarkTest(int Instruction, DWORD Round, DWORD *pTime)
{
	if (m_pCasManager == nullptr)
		return E_UNEXPECTED;

	if (!m_pCasManager->DescrambleBenchmarkTest(Instruction, Round, pTime))
		return E_FAIL;

	return S_OK;
}


HRESULT CCasProcessor::OnError(HRESULT hr, LPCWSTR pszText, LPCWSTR pszAdvise, LPCWSTR pszSystemMessage)
{
	if (m_pClient == nullptr)
		return E_UNEXPECTED;

	TVTest::Interface::ErrorInfo Info;

	Info.hr = hr;
	Info.pszText = pszText;
	Info.pszAdvise = pszAdvise;
	Info.pszSystemMessage = pszSystemMessage;

	return m_pClient->OnError(&Info);
}


void CCasProcessor::OutLog(TVTest::Interface::LogType Type, LPCWSTR pszFormat, ...)
{
	if (m_pClient != nullptr) {
		va_list Args;
		WCHAR szBuffer[256];

		va_start(Args, pszFormat);
		::wvnsprintfW(szBuffer, _countof(szBuffer), pszFormat, Args);
		va_end(Args);
		m_pClient->OutLog(Type, szBuffer);
	}
}


void CCasProcessor::OnCasEvent(UINT Event, void *pParam)
{
	if (m_pClient == nullptr)
		return;

	switch (Event) {
	case TVCAS::EVENT_EMM_PROCESSED:
		OutLog(TVTest::Interface::LOG_INFO, L"EMM処理を行いました。");
		break;

	case TVCAS::EVENT_EMM_ERROR:
		{
			const TVCAS::EmmErrorInfo *pInfo = static_cast<TVCAS::EmmErrorInfo*>(pParam);

			if (pInfo->pszText != nullptr && pInfo->pszText[0] != L'\0')
				OutLog(TVTest::Interface::LOG_ERROR, L"EMM処理でエラーが発生しました。(%s)", pInfo->pszText);
			else
				OutLog(TVTest::Interface::LOG_ERROR, L"EMM処理でエラーが発生しました。");
		}
		break;

	case TVCAS::EVENT_ECM_ERROR:
		if (m_pCasManager != nullptr) {
			const TVCAS::EcmErrorInfo *pInfo = static_cast<TVCAS::EcmErrorInfo*>(pParam);

			if (m_ActiveServiceID != 0
				&& m_pCasManager->GetEcmPIDByServiceID(m_ActiveServiceID) == pInfo->EcmPID) {
				m_pClient->Notify(TVTest::Interface::NOTIFY_ERROR, L"スクランブル解除でエラーが発生しました");

				if (pInfo->pszText != nullptr && pInfo->pszText[0] != L'\0')
					OutLog(TVTest::Interface::LOG_ERROR, L"ECM処理でエラーが発生しました。(%s)", pInfo->pszText);
				else
					OutLog(TVTest::Interface::LOG_ERROR, L"ECM処理でエラーが発生しました。");
			}
		}
		break;

	case TVCAS::EVENT_ECM_REFUSED:
		if (m_pCasManager != nullptr) {
			const TVCAS::EcmErrorInfo *pInfo = static_cast<TVCAS::EcmErrorInfo*>(pParam);

			if (m_ActiveServiceID != 0
				&& m_pCasManager->GetEcmPIDByServiceID(m_ActiveServiceID) == pInfo->EcmPID) {
				m_pClient->Notify(TVTest::Interface::NOTIFY_WARNING, L"契約されていないため視聴できません");
			}
		}
		break;

	case TVCAS::EVENT_CARD_READER_HUNG:
		m_pClient->Notify(TVTest::Interface::NOTIFY_ERROR, L"カードリーダーから応答がありません");
		OutLog(TVTest::Interface::LOG_ERROR, L"カードリーダーから応答がありません。");
		break;
	}
}




CCasProcessor::CCasClient::CCasClient(CCasProcessor *pCasProcessor)
	: m_pCasProcessor(pCasProcessor)
{
	m_pCasProcessor->AddRef();
}


CCasProcessor::CCasClient::~CCasClient()
{
	m_pCasProcessor->Release();
}


LRESULT CCasProcessor::CCasClient::OnEvent(UINT Event, void *pParam)
{
	m_pCasProcessor->OnCasEvent(Event, pParam);
	return 0;
}


LRESULT CCasProcessor::CCasClient::OnError(const TVCAS::ErrorInfo *pInfo)
{
	return m_pCasProcessor->OnError(E_FAIL, pInfo->pszText, pInfo->pszAdvise, pInfo->pszSystemMessage);
}


void CCasProcessor::CCasClient::OutLog(TVCAS::LogType Type, LPCWSTR pszMessage)
{
	m_pCasProcessor->OutLog((TVTest::Interface::LogType)Type, L"%s", pszMessage);
}




class CCasProcessorPlugin : public TVTest::CTVTestPlugin
{
public:
	bool GetPluginInfo(TVTest::PluginInfo *pInfo) override;
	bool Initialize() override;
	bool Finalize() override;
};


bool CCasProcessorPlugin::GetPluginInfo(TVTest::PluginInfo *pInfo)
{
	pInfo->Type           = TVTest::PLUGIN_TYPE_NORMAL | TVTest::PLUGIN_FLAG_NOUNLOAD;
	pInfo->Flags          = 0;
	pInfo->pszPluginName  = L"CAS Processor";
	pInfo->pszCopyright   = L"Public Domain";
	pInfo->pszDescription = L"CAS処理を行います。";
	return true;
}


bool CCasProcessorPlugin::Initialize()
{
	TVTest::TSProcessorInfo Info;
	Info.Size            = sizeof(Info);
	Info.Flags           = 0;
	Info.pTSProcessor    = new CCasProcessor;
	Info.ConnectPosition = TVTest::TS_RPOCESSOR_CONNECT_POSITION_VIEWER;
	bool fResult = m_pApp->RegisterTSProcessor(&Info);
	Info.pTSProcessor->Release();
	if (!fResult) {
		m_pApp->AddLog(L"TSプロセッサーを登録できません。");
		return false;
	}

	return true;
}


bool CCasProcessorPlugin::Finalize()
{
	return true;
}




TVTest::CTVTestPlugin *CreatePluginClass()
{
	return new CCasProcessorPlugin;
}
