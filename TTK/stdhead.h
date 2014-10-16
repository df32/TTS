#pragma once

#include "sapi.h"
#include "sphelper.h"
#include <windows.h>
#include <iostream>

#define nil 0
#define Widden Widen

using namespace std;

struct SpBox {
	ISpVoice*			VoiceObj;	//语音对象
	ISpObjectToken*		Token;		//语音选择的对象
	wstring				Description;//文字描述
	SPVOICESTATUS		Status;		//语音状态

	wstring				Text;		//朗读的文本
	DWORD				Flags;		//朗读模式

	WCHAR*				Lang;		//语言代码


	HRESULT Speak() {				//开始朗读
		return VoiceObj->Speak(Text.c_str(), Flags, 0);
	}

	HRESULT Pause() {				//暂停
		return VoiceObj->Pause();
	}

	HRESULT Resume() {				//继续
		return VoiceObj->Resume();
	}

	ULONG	GetStatus() {			//返回当前语音状态: 0完毕;SPRS_IS_SPEAKING 正在处理
		VoiceObj->GetStatus(&Status, 0);
		return Status.dwRunningState;	
	}	


	HRESULT Initialize() {			//注册
		HRESULT hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL, IID_ISpVoice,
			(void **)&VoiceObj);
		if (SUCCEEDED(hr))
			return VoiceObj->GetVoice(&Token);
		return hr;
	}

	HRESULT WaitForDone(ULONG outime = INFINITE) {			//等待结束
		return VoiceObj->WaitUntilDone(outime);
	}

	HRESULT GetLanguageFromToken(){	//获取语言信息
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
ISpStream*				pWavStream;			//wav流	（音频接口）
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





