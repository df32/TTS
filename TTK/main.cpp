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
			dfBox.Flags |= SPF_PURGEBEFORESPEAK;	//����ǰ�������������������
		if (argstr.find("/f") != string::npos)
			dfBox.Flags |= SPF_IS_FILENAME;			//�ʶ��ĵ������ĵ���
		if (argstr.find("/x") != string::npos)
			dfBox.Flags |= SPF_IS_XML;				//���ݴ���XML��ʽ	
		if (argstr.find("/nx") != string::npos)
			dfBox.Flags |= SPF_IS_NOT_XML;			//����û��XML��ʽ	
		if (argstr.find("/px") != string::npos)
			dfBox.Flags |= SPF_PERSIST_XML;			//Global state changes in the XML markup will persist across speak calls. 
		if (argstr.find("/sp") != string::npos)	
			dfBox.Flags |= SPF_NLP_SPEAK_PUNC;		//�ʶ����
		
		dfBox.Text = Widen(argv[1]);
		dfBox.Speak();
		break;
	}


	dfBox.WaitForDone(3000);
	if (dfBox.GetStatus() == SPRS_IS_SPEAKING) {
		cout << "��ǰ������\t(Y/N)" << endl
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

	IEnumSpObjectTokens*    pEnum;				//������ѡ��
	ISpObjectToken*			pVoiceToken;		//��ǰ����ѡ��
	ULONG					ulCount = 0;		//ѡ�����
	WCHAR*					pVoiceText;			//����˵������

	hr = SpEnumTokens(SPCAT_VOICES, NULL, NULL, &pEnum);
	if (SUCCEEDED(hr)){
		hr = pEnum->GetCount(&ulCount);
	}

	cout << "��ӡ���������б�...(" << ulCount << ")" << endl
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
	//����һ����������󶨵�wav�ļ�
	hr = pVoice->GetOutputStream(&cpOldStream);		//��ȡ��ǰ�����
	if (SUCCEEDED(hr))
		hr = OriginalFmt.AssignFormat(cpOldStream);
	else hr = E_FAIL;
	if (SUCCEEDED(hr))
		hr = SPBindToFile((LPCWSTR)&fname, SPFM_CREATE_ALWAYS, &pWavStream, //����Ƶ���󶨵�ָ���ļ�
		&OriginalFmt.FormatId(), OriginalFmt.WaveFormatExPtr());
	if (SUCCEEDED(hr))
		pVoice->SetOutput(pWavStream, 1);	//�����������
	
	if (!SUCCEEDED(hr))
		return -1;
	return 0;
}

int overputStream(){
	// �ȴ��ʶ�����
	pVoice->WaitUntilDone(INFINITE);
	pWavStream->Release();

	//��������¶�λ��ԭ������
	pVoice->SetOutput(cpOldStream,0);
	return 0;
}*/

