#include "stdhead.h"

SpBox dfBox;


int main(int argc, char* argv[])
{	
	for (int i = 0; i < argc; i++) 	{
		cout << argv[i]<< "\t";
	}
	cout << endl << endl;
	if (argc == 1) 	{
		printHelp();
		return 0;
	};

	Initialize();
	dfBox.Flags = SPF_ASYNC;
	switch (argc){
	case 2: 
		if (!_stricmp(argv[1], "/?")){
			printHelp(1);
			break;
		}
		dfBox.Text = Widen(argv[1]);
		dfBox.Speak();
		break;
	case 3:
		if (!_stricmp(argv[1], "/f")){
			
			dfBox.Flags |= SPF_IS_FILENAME;
			dfBox.Text = Widen(argv[1]);
			dfBox.Speak();
			break;
		}
	default:
		string argstr;
		for (int i = 2; i < argc; i++){
			argstr.append(_strlwr(argv[i]));
		}
		
		if (argstr.find("/p") != string::npos)
			dfBox.Flags |= SPF_PURGEBEFORESPEAK;	//播放前清除所有其他语音请求
		if (argstr.find("/f") != string::npos)
			dfBox.Flags |= SPF_IS_FILENAME;			//朗读文档而非文档名
		if (argstr.find("/x") != string::npos)
			dfBox.Flags |= SPF_IS_XML;				//内容带有XML格式	
		if (argstr.find("/nx") != string::npos)
			dfBox.Flags |= SPF_IS_NOT_XML;			//内容没有XML格式	
		if (argstr.find("/px") != string::npos)
			dfBox.Flags |= SPF_PERSIST_XML;			//Global state changes in the XML markup will persist across speak calls. 
		if (argstr.find("/sp") != string::npos)	
			dfBox.Flags |= SPF_NLP_SPEAK_PUNC;		//朗读标点
		
		dfBox.Text = Widen(argv[1]);
		dfBox.Speak();
		break;
	}


	dfBox.WaitForDone(3000);
	if (dfBox.GetStatus() == SPRS_IS_SPEAKING) {
		cout << "提前结束？\t(Y/N)" << endl
			<< "Quit?\t\t(Y/N)" << endl;
		char cmd[5];

		while (dfBox.GetStatus() == SPRS_IS_SPEAKING && _stricmp(cmd, "y"))
			cin >> cmd;
	}

	Finalize();
	return 0;
}

int Initialize(){
	if (FAILED(::CoInitialize(nil)))
		return -1;
	return (int)SUCCEEDED(dfBox.Initialize());
}

void Finalize(){
	dfBox.Finalize();
	::CoUninitialize();
}


int printHelp(int type){
	cout << "This is print help" << endl;

	if (type) printList();
	return 0;
}


int printList(){

	IEnumSpObjectTokens*    pEnum;				//语音总选择
	ISpObjectToken*			pVoiceToken;		//当前语音选择
	ULONG					ulCount = 0;		//选择个数
	WCHAR*					pVoiceText;			//语音说明文字

	hr = SpEnumTokens(SPCAT_VOICES, NULL, NULL, &pEnum);
	if (SUCCEEDED(hr)){
		hr = pEnum->GetCount(&ulCount);
	}

	cout << "打印所有语音列表...(" << ulCount << ")" << endl
		<< "Print the list of all voices...(" << ulCount << ")" << endl;

	while (SUCCEEDED(hr)&&ulCount--){
		
		if (SUCCEEDED(hr))
			pEnum->Next(1, &pVoiceToken, 0);
		if (SUCCEEDED(hr))
			hr = SpGetDescription(pVoiceToken, (WCHAR**)&pVoiceText, 0);
		if (SUCCEEDED(hr)){
			string str = Narrow(pVoiceText);
			cout << "\t" << str.c_str() << endl;
		}
	}
	if (!SUCCEEDED(hr))
		return -1;
	return 0;
}

/*
int outputStream(wstring fname){
	//创建一个输出流，绑定到wav文件
	hr = pVoice->GetOutputStream(&cpOldStream);		//获取当前输出流
	if (SUCCEEDED(hr))
		hr = OriginalFmt.AssignFormat(cpOldStream);
	else hr = E_FAIL;
	if (SUCCEEDED(hr))
		hr = SPBindToFile((LPCWSTR)&fname, SPFM_CREATE_ALWAYS, &pWavStream, //将音频流绑定到指定文件
		&OriginalFmt.FormatId(), OriginalFmt.WaveFormatExPtr());
	if (SUCCEEDED(hr))
		pVoice->SetOutput(pWavStream, 1);	//设置输出对象
	
	if (!SUCCEEDED(hr))
		return -1;
	return 0;
}

int overputStream(){
	// 等待朗读结束
	pVoice->WaitUntilDone(INFINITE);
	pWavStream->Release();

	//把输出重新定位到原来的流
	pVoice->SetOutput(cpOldStream,0);
	return 0;
}*/

