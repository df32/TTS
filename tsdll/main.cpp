#include "stdhead.h"

#define _CRT_SECURE_NO_WARNINGS 1

using namespace std;


HRESULT		g_hr = -1;				//ȫ�ֳ�ʼ���Ƿ�ɹ�
int			g_Mcount = 0;			//Measureʵ������


bool HRtestify(HRESULT hr, const WCHAR* append);
__inline bool HRtestify(HRESULT hr) { return HRtestify(hr, L""); }


struct SpBox 
{
	ISpVoice*			VoiceObj;		//��������
	ISpObjectToken*		Token;			//����ѡ��Ķ���
	wstring				Description;	//��������
	SPVOICESTATUS		Status;			//����״̬

	wstring				Text;			//�ʶ����ı�
	DWORD				Flags;			//�ʶ�ģʽ
	bool				Paused = false;	//�Ƿ���ͣ

	//WCHAR*				Lang;		//���Դ���

	void*				rm;				//rmָ��
	bool				valid = false;			//�Ƿ�ע��ɹ�
	bool				Initialized = false;	//�Ƿ��Ѿ���ȡMeasure����

	HRESULT Speak()				//��ʼ�ʶ�
	{
		UnPause();
		return VoiceObj->Speak(Text.c_str(), Flags, 0);
	}

	HRESULT Speak(LPCWSTR pwcs)
	{
		UnPause();
		return VoiceObj->Speak(pwcs, Flags, 0);
	}


	HRESULT Pause()				//��ͣ�ʶ�
	{
		if (!Paused)
		{
			Paused = true;
			return VoiceObj->Pause();
		}
		return 0;
	}

	HRESULT UnPause()			//�����ʶ�
	{
		if (Paused)
		{
			Paused = false;
			return VoiceObj->Resume();
		}
		return 0;
	}

	HRESULT TogglePause()		//�л���ͣ
	{
		
		Paused = !Paused;

		if (Paused)
			return VoiceObj->Pause();
		else
			return VoiceObj->Resume();
	}

	HRESULT SetVolume(USHORT usVolume)	//�����������ʶ���ʼǰ��Ч��
	{
		return VoiceObj->SetVolume(usVolume);
	}

	HRESULT SetRate(long RateAdjust)	//�������٣�ͬ�ϣ�
	{
		return VoiceObj->SetRate(RateAdjust);
	}


	ULONG	GetStatus() 			//���ص�ǰ����״̬: 0���;SPRS_IS_SPEAKING ���ڴ���
	{
		VoiceObj->GetStatus(&Status, 0);
		return Status.dwRunningState;
	}


	/*HRESULT Initialize() {			//ע��
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


	/*HRESULT WaitForDone(ULONG outime = INFINITE) 		//�ȴ�����
	{
		return VoiceObj->WaitUntilDone(outime);
	}*/

	/*HRESULT GetLanguageFromToken()						//��ȡ������Ϣ
	{
		ISpDataKey* cpKey;
		HRESULT hr = Token->OpenKey(L"Attributes", &cpKey);
		if (FAILED(hr)) return hr;
		hr = cpKey->GetStringValue(L"Language", &Lang);
		return hr;
	}*/

	bool RequireLanguage(LPCWSTR langid);				//����ָ�����Ե�����

	SpBox();
	~SpBox();
};


//���캯��
SpBox::SpBox()
{
	//ȫ�ֳ�ʼ��
	if (g_Mcount < 1 || FAILED(g_hr))
		g_hr = ::CoInitialize(0);

	if FAILED(g_hr) 
	{
		RmLog(LOG_ERROR, L"T2S.dll failed to coinitialize. Make sure your Windows Narrator is available!");
		return;
	}

	//ע��
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


//��ȡָ��������LangId
HRESULT SpGetLanguage(ISpObjectToken * pObjToken, PWSTR *ppszLangId)
{
	ISpDataKey* cpKey;

	HRESULT hr = pObjToken->OpenKey(L"Attributes", &cpKey);

	if (FAILED(hr)) return hr;

	hr = cpKey->GetStringValue(L"Language", ppszLangId);

	return hr;
}


//����ָ�����Ե�����
bool SpBox::RequireLanguage(LPCWSTR langid)
{
	IEnumSpObjectTokens*    pEnum;				//������ѡ��
	ISpObjectToken*			pVoiceToken;		//��ǰ����ѡ��
	ULONG					ulCount = 0;		//ѡ�����
	WCHAR*					pVoiceLang;			//����LangId
	WCHAR*					pVoiceText;			//����˵������

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


//��������
SpBox::~SpBox()
{
	g_Mcount--;

	//�ͷ�
	if (Token != nullptr)
		Token->Release();
	if (VoiceObj != nullptr)
		VoiceObj->Release();

	//ȫ��ע��
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

	Measure->Text	= RmReadString(rm, L"Text", L"");			//�ʶ��ı�
	LPCWSTR FileName= RmReadPath(rm, L"FileName", L""); 		//�ʶ��ļ�
	wstring Flags	= RmReadString(rm, L"Flags", L"");			//Flag����
	//wstring LangID	= RmReadString(rm, L"LangId", L"");			//��������
	int Volume		= RmReadInt(rm, L"Volume", 0);				//����
	long Rate		= RmReadInt(rm, L"Rate", 0);				//����
	
	
	Measure->Flags = SPF_ASYNC;
	for (wstring::size_type i = 0; i < Flags.size(); i++)
	{
		Flags[i] = towupper(Flags[i]);
	}

	if (!Flags.empty()) {
		if (Flags.find(L"PURGEBEFORESPEAK") != wstring::npos)
			Measure->Flags |= SPF_PURGEBEFORESPEAK;	//����ǰ�������������������

		if (Flags.find(L"NLP_SPEAK_PUNC") != wstring::npos)
			Measure->Flags |= SPF_NLP_SPEAK_PUNC;	//�ʶ����

		if (Flags.find(L"IS_XML") != wstring::npos)
			Measure->Flags |= SPF_IS_XML;			//���ݴ���XML��ʽ

		if (Flags.find(L"IS_NOT_XML") != wstring::npos)
			Measure->Flags |= SPF_IS_NOT_XML;		//����û��XML��ʽ

		if (Flags.find(L"IS_FILENAME") != wstring::npos)
		{
			Measure->Flags |= SPF_IS_FILENAME;		//�ʶ��ĵ������ĵ���
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

	//��֧�ֶ�̬�����Ĳ���
	if (!Measure->Initialized)
	{
		wstring LangID = RmReadString(rm, L"LangId", L"");			//��������
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
