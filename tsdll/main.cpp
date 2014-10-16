#include "stdhead.h"

#define _CRT_SECURE_NO_WARNINGS 1

using namespace std;


HRESULT		g_hr = -1;				//全局初始化是否成功
int			g_Mcount = 0;			//Measure实例个数


bool HRtestify(HRESULT hr, const WCHAR* append);
__inline bool HRtestify(HRESULT hr) { return HRtestify(hr, L""); }


struct SpBox 
{
	ISpVoice*			VoiceObj;		//语音对象
	ISpObjectToken*		Token;			//语音选择的对象
	wstring				Description;	//文字描述
	SPVOICESTATUS		Status;			//语音状态

	wstring				Text;			//朗读的文本
	DWORD				Flags;			//朗读模式
	bool				Paused = false;	//是否暂停

	//WCHAR*				Lang;		//语言代码

	void*				rm;				//rm指针
	bool				valid = false;			//是否注册成功
	bool				Initialized = false;	//是否已经读取Measure设置

	HRESULT Speak()				//开始朗读
	{
		UnPause();
		return VoiceObj->Speak(Text.c_str(), Flags, 0);
	}

	HRESULT Speak(LPCWSTR pwcs)
	{
		UnPause();
		return VoiceObj->Speak(pwcs, Flags, 0);
	}


	HRESULT Pause()				//暂停朗读
	{
		if (!Paused)
		{
			Paused = true;
			return VoiceObj->Pause();
		}
		return 0;
	}

	HRESULT UnPause()			//继续朗读
	{
		if (Paused)
		{
			Paused = false;
			return VoiceObj->Resume();
		}
		return 0;
	}

	HRESULT TogglePause()		//切换暂停
	{
		
		Paused = !Paused;

		if (Paused)
			return VoiceObj->Pause();
		else
			return VoiceObj->Resume();
	}

	HRESULT SetVolume(USHORT usVolume)	//设置音量（朗读开始前有效）
	{
		return VoiceObj->SetVolume(usVolume);
	}

	HRESULT SetRate(long RateAdjust)	//设置语速（同上）
	{
		return VoiceObj->SetRate(RateAdjust);
	}


	ULONG	GetStatus() 			//返回当前语音状态: 0完毕;SPRS_IS_SPEAKING 正在处理
	{
		VoiceObj->GetStatus(&Status, 0);
		return Status.dwRunningState;
	}


	/*HRESULT Initialize() {			//注册
		HRESULT hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice,
			(void **)&VoiceObj);
		if (SUCCEEDED(hr))
			return VoiceObj->GetVoice(&Token);
		return hr;
	}*/


	/*void Finalize() {
	if (Token) Token->Release();
	VoiceObj->Release();
	}*/


	/*HRESULT WaitForDone(ULONG outime = INFINITE) 		//等待结束
	{
		return VoiceObj->WaitUntilDone(outime);
	}*/

	/*HRESULT GetLanguageFromToken()						//获取语言信息
	{
		ISpDataKey* cpKey;
		HRESULT hr = Token->OpenKey(L"Attributes", &cpKey);
		if (FAILED(hr)) return hr;
		hr = cpKey->GetStringValue(L"Language", &Lang);
		return hr;
	}*/

	bool RequireLanguage(LPCWSTR langid);				//请求指定语言的语音

	SpBox();
	~SpBox();
};


//构造函数
SpBox::SpBox()
{
	//全局初始化
	if (g_Mcount < 1 || FAILED(g_hr))
		g_hr = ::CoInitialize(0);

	if FAILED(g_hr) 
	{
		RmLog(LOG_ERROR, L"T2S.dll failed to coinitialize. Make sure your Windows Narrator is available!");
		return;
	}

	//注册
	HRESULT hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice, (void **)&VoiceObj);

	if (SUCCEEDED(hr))
		hr = VoiceObj->GetVoice(&Token);
	
	if (SUCCEEDED(hr))
	{
		WCHAR* VoiceText;
		hr = SpGetDescription(Token, &VoiceText, 0);
		Description = VoiceText;
	}	

	else 
	{ 
		HRtestify(hr, L" in Initialize");
	}

	valid = SUCCEEDED(hr) ? true : false;
	g_Mcount++;

}


//获取指定语音的LangId
HRESULT SpGetLanguage(ISpObjectToken * pObjToken, PWSTR *ppszLangId)
{
	ISpDataKey* cpKey;

	HRESULT hr = pObjToken->OpenKey(L"Attributes", &cpKey);

	if (FAILED(hr)) return hr;

	hr = cpKey->GetStringValue(L"Language", ppszLangId);

	return hr;
}


//请求指定语言的语音
bool SpBox::RequireLanguage(LPCWSTR langid)
{
	IEnumSpObjectTokens*    pEnum;				//语音总选择
	ISpObjectToken*			pVoiceToken;		//当前语音选择
	ULONG					ulCount = 0;		//选择个数
	WCHAR*					pVoiceLang;			//语音LangId
	WCHAR*					pVoiceText;			//语音说明文字

	HRESULT hr = SpEnumTokens(SPCAT_VOICES, NULL, NULL, &pEnum);
	if (SUCCEEDED(hr))
	{
		hr = pEnum->GetCount(&ulCount);
	}

	while (SUCCEEDED(hr) && ulCount--){

		if (SUCCEEDED(hr))
			pEnum->Next(1, &pVoiceToken, 0);
		
		if (SUCCEEDED(hr))
			hr = SpGetLanguage(pVoiceToken, &pVoiceLang);

		if (SUCCEEDED(hr) && !_wcsnicmp(pVoiceLang, langid, 3))
		{
			Token = pVoiceToken;
			VoiceObj->SetVoice(Token);
			SpGetDescription(Token, &pVoiceText, 0);
			Description = pVoiceText;
			return true;
		}
	}
	return false;
}


//析构函数
SpBox::~SpBox()
{
	g_Mcount--;

	//释放
	if (Token != nullptr)
		Token->Release();
	if (VoiceObj != nullptr)
		VoiceObj->Release();

	//全局注销
	if (g_Mcount < 1)
		::CoUninitialize();
}



PLUGIN_EXPORT void Initialize(void** data, void* rm)
{
	SpBox* Measure = new SpBox;
	*data = Measure;

	if (!Measure->valid)
	{
		wstring log = L"T2S.dll: [";
		log += RmGetMeasureName(rm);
		log += L"] is invalid!";
		RmLog(LOG_ERROR, log.c_str());
	}

	Measure->rm = rm;

}


PLUGIN_EXPORT void Finalize(void* data)
{
	SpBox* Measure = (SpBox*)data;
	delete Measure;
}

PLUGIN_EXPORT void Reload(void* data, void* rm, double* maxValue)
{
	SpBox* Measure = (SpBox*)data;
	Measure->rm = rm;

	Measure->Text	= RmReadString(rm, L"Text", L"");			//朗读文本
	LPCWSTR FileName= RmReadPath(rm, L"FileName", L""); 		//朗读文件
	wstring Flags	= RmReadString(rm, L"Flags", L"");			//Flag参数
	//wstring LangID	= RmReadString(rm, L"LangId", L"");			//语音语言
	int Volume		= RmReadInt(rm, L"Volume", 0);				//音量
	long Rate		= RmReadInt(rm, L"Rate", 0);				//语速
	
	
	Measure->Flags = SPF_ASYNC;
	for (wstring::size_type i = 0; i < Flags.size(); i++)
	{
		Flags[i] = towupper(Flags[i]);
	}

	if (!Flags.empty()) {
		if (Flags.find(L"PURGEBEFORESPEAK") != wstring::npos)
			Measure->Flags |= SPF_PURGEBEFORESPEAK;	//播放前清除所有其他语音请求

		if (Flags.find(L"NLP_SPEAK_PUNC") != wstring::npos)
			Measure->Flags |= SPF_NLP_SPEAK_PUNC;	//朗读标点

		if (Flags.find(L"IS_XML") != wstring::npos)
			Measure->Flags |= SPF_IS_XML;			//内容带有XML格式

		if (Flags.find(L"IS_NOT_XML") != wstring::npos)
			Measure->Flags |= SPF_IS_NOT_XML;		//内容没有XML格式

		if (Flags.find(L"IS_FILENAME") != wstring::npos)
		{
			Measure->Flags |= SPF_IS_FILENAME;		//朗读文档而非文档名
			Measure->Text = FileName;
		}
	}
	else if (wcslen(FileName) > 0)
	{
		Measure->Flags |= SPF_IS_FILENAME;
		Measure->Text = FileName;
	}

	if (Volume > 0)
		Measure->SetVolume(Volume);

	if (Rate)
		Measure->SetRate(Rate);

	//不支持动态变量的部分
	if (!Measure->Initialized)
	{
		wstring LangID = RmReadString(rm, L"LangId", L"");			//语音语言
		if (!LangID.empty())
		{
			if (!Measure->RequireLanguage(LangID.c_str()))
				RmLog(LOG_ERROR, L"T2S.dll: Mismatched LangId");
		}
		if (!Measure->Text.empty() && RmReadInt(rm, L"SpeakOnLoad", 0) > 0)
		{
			HRtestify(Measure->Speak(), L" in Speak()");
		}

		Measure->Initialized = true;

	}
}

PLUGIN_EXPORT double Update(void* data)
{
	SpBox* Measure = (SpBox*)data;
	return int (Measure->GetStatus() > 0);
}

PLUGIN_EXPORT LPCWSTR GetString(void* data)
{
	SpBox* Measure = (SpBox*)data;
	return Measure->Description.c_str();
}

PLUGIN_EXPORT void ExecuteBang(void* data, LPCWSTR args) 
{
	SpBox* Measure = (SpBox*)data;

	if (!_wcsnicmp(args, L"_Reload",7))
	{
		Measure->Initialized = false;
		Reload(data, Measure->rm, 0);
	}
	else
	{
		if (!Measure->valid) return;

		if (!_wcsicmp(args, L"_Speak"))
			HRtestify(Measure->Speak(), L" in Speak()");

		else if (!_wcsicmp(args, L"_Pause"))
			HRtestify(Measure->Pause(), L" in Pause()");

		else if (!_wcsicmp(args, L"_UnPause"))
			HRtestify(Measure->UnPause(), L" in UnPause()");

		else if (wcslen(args) > 1)
			HRtestify(Measure->Speak(args), L" in Speak2()");
	}

}

bool HRtestify(HRESULT hr, LPCWSTR append)
{
	if (SUCCEEDED(hr))
		return true;

	wstring log = L"T2S.dll: Critical error";
	log += append;
	log += L"! Error Code: ";
	WCHAR ErrorCode[15];
	_ltow_s(hr, ErrorCode, 16);
	log += ErrorCode;

	RmLog(LOG_ERROR, log.c_str());
	return false;
}
