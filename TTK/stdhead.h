#pragma once

#include "sapi.h"
#include "sphelper.h"
#include <windows.h>
#include <iostream>

#define nil 0
#define Widden Widen

using namespace std;

struct SpBox {
	ISpVoice*			VoiceObj;	//��������
	ISpObjectToken*		Token;		//����ѡ��Ķ���
	wstring				Description;//��������
	SPVOICESTATUS		Status;		//����״̬

	wstring				Text;		//�ʶ����ı�
	DWORD				Flags;		//�ʶ�ģʽ

	WCHAR*				Lang;		//���Դ���


	HRESULT Speak() {				//��ʼ�ʶ�
		return VoiceObj->Speak(Text.c_str(), Flags, 0);
	}

	HRESULT Pause() {				//��ͣ
		return VoiceObj->Pause();
	}

	HRESULT Resume() {				//����
		return VoiceObj->Resume();
	}

	ULONG	GetStatus() {			//���ص�ǰ����״̬: 0���;SPRS_IS_SPEAKING ���ڴ���
		VoiceObj->GetStatus(&Status, 0);
		return Status.dwRunningState;	
	}	


	HRESULT Initialize() {			//ע��
		HRESULT hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice,
			(void **)&VoiceObj);
		if (SUCCEEDED(hr))
			return VoiceObj->GetVoice(&Token);
		return hr;
	}

	HRESULT WaitForDone(ULONG outime = INFINITE) {			//�ȴ�����
		return VoiceObj->WaitUntilDone(outime);
	}

	HRESULT GetLanguageFromToken(){	//��ȡ������Ϣ
		ISpDataKey* cpKey;
		WCHAR*		Lang;
		HRESULT hr = Token->OpenKey(L"Attributes", &cpKey);
		if (FAILED(hr)) return hr;
		hr = cpKey->GetStringValue(L"Language", &Lang);
		return hr;
	}

	void Finalize() {
		if (Token) Token->Release();
		VoiceObj->Release();
	}
};



HRESULT					hr;
/*
CSpStreamFormat			OriginalFmt;
ISpStream*				pWavStream;			//wav��	����Ƶ�ӿڣ�
ISpStreamFormat*		cpOldStream;
*/

//		Copyright(C) 2013 Birunthan Mohanathas
std::wstring Widen(const char* str, int strLen = -1, int cp = CP_ACP);
std::wstring Widen(const std::string& str, int cp = CP_ACP) { return Widen(str.c_str(), (int)str.length(), cp); }

std::string Narrow(const WCHAR* str, int strLen = -1, int cp = CP_ACP);
std::string Narrow(const std::wstring& str, int cp = CP_ACP) { return Narrow(str.c_str(), (int)str.length(), cp); }

std::wstring Widen(const char* str, int strLen, int cp){

	std::wstring wideStr;

	if (str && *str)
	{
		if (strLen == -1)
		{
			strLen = strlen(str);
		}

		int bufLen = MultiByteToWideChar(cp, 0, str, strLen, nullptr, 0);
		if (bufLen > 0)
		{
			wideStr.resize(bufLen);
			MultiByteToWideChar(cp, 0, str, strLen, &wideStr[0], bufLen);
		}
	}
	return wideStr;
}

std::string Narrow(const WCHAR* str, int strLen, int cp)
{
	std::string narrowStr;

	if (str && *str)
	{
		if (strLen == -1)
		{
			strLen = (int)wcslen(str);
		}

		int bufLen = WideCharToMultiByte(cp, 0, str, strLen, nullptr, 0, nullptr, nullptr);
		if (bufLen > 0)
		{
			narrowStr.resize(bufLen);
			WideCharToMultiByte(cp, 0, str, strLen, &narrowStr[0], bufLen, nullptr, nullptr);
		}
	}
	return narrowStr;
}
int main(int argc, char* argv[]);
int Initialize();
void Finalize();
int printHelp(int type = 0);
int printList();
// int outputStream(wstring fname);





