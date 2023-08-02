#include "windows.h"
#include <string>
#include <iostream>
#include "getopt.h"
#include "sapi.h"
#include "sphelper.h"
#pragma comment(lib, "sapi.lib")


using namespace std;

typedef struct
{
	char text[512]; // �ʶ��ı�
	int  volume;    // �ʶ���������Ч��Χ[0,100]
	int  rate;      // �ʶ����ʣ���Ч��Χ[-10,10]
	int  forTime;   // �ʶ�ѭ������

}VOICE_OPTS;

static void lib_tts_usage(const char *program)
{
	cout << endl;
	cout << program << " [-v volume] [-r rate] [-f forTime] {-t text}\n" << endl;
	cout << "������ʹ�øù��߿��Խ��ı�ת��������" << endl;
	cout << "�����б�\n"
			"  -v --volume   �ʶ�����(int)����Ч��Χ[0,100]\n" 
			"  -r --rate     �ʶ�����(int)����Ч��Χ[-10,10]\n"
			"  -t --text     �ʶ��ı�(string)��֧����Ӣ���\n"
			"  -? --help     �����ĵ�\n"
			"  -f --forTime  �ʶ�ѭ������(int)"<< endl;
}

wchar_t * char_to_wchar(const char *str)
{
	wchar_t *p_wchar = NULL;
	int len = MultiByteToWideChar(CP_ACP, 0, str, -1, NULL, 0);
	p_wchar = new wchar_t[len+1];
	wmemset(p_wchar, 0, len + 1);
	MultiByteToWideChar(CP_ACP, 0, str, -1, p_wchar, len);
	return p_wchar;
}

static void lib_tts_opt_init(VOICE_OPTS *voice_opts)
{
	voice_opts->volume = 80;
	voice_opts->rate = 0;
	voice_opts->forTime = 1;
}

static int lib_tts_opts_get(int argc, char *argv[], const struct option *long_options, VOICE_OPTS *voice_opts)
{
	int opt = 0, val = 0;

	while ((opt = getopt_long(argc, argv, "v:f:r:t:?", long_options, NULL)) != EOF)
	{
		switch (opt)
		{
		case 'v':
			val = atoi(optarg);
			if (val < 0 || val > 100)
			{
				cout << "error: volume's valid range [0, 100]" << endl;
				return 1;
			}
			voice_opts->volume = val;
			break;
		case 'f':
			val = atoi(optarg);
			if (val < 0)
			{
				cout << "error: volume's val more than 0" << endl;
				return 1;
			}
			voice_opts->forTime = atoi(optarg);
			break;
		case 'r':
			val = atoi(optarg);
			if (val < -10 || val > 10)
			{
				cout << "error: rate's valid range [-10, 10]" << endl;
				return 1;
			}
			voice_opts->rate = atoi(optarg);
			break;
		case 't':
			snprintf(voice_opts->text, sizeof(voice_opts->text), "%s", optarg);
			break;
		case '?':
		default:
			lib_tts_usage(argv[0]);
			return 1;
		}
	}
	if (strlen(voice_opts->text) == 0)
	{
		cout << "error: --text argument isn't empty" << endl;
		return 1;
	}
	return 0;
}


int main(int argc, char *argv[])
{
	int rc = 0;
	VOICE_OPTS voice_opts;
	memset(&voice_opts, 0, sizeof(VOICE_OPTS));
	/* VOICE_OPTS ������ʼ�� */
	lib_tts_opt_init(&voice_opts);
	/* �����в�����ȡ */
	static struct option long_options[] = {
	{"help",    0, 0,  '?'},
	{"volume",  1, 0,  'v'},
	{"rate",    1, 0,  'r'},
	{"forTime", 1, 0,  'f'},
	{"text",    1, 0,  't'},
	{0 , 0, 0, 0}
	};
	if ((rc = lib_tts_opts_get(argc, argv, long_options, &voice_opts)) != 0)
	{
		return -1;
	}
	
	const char *str = voice_opts.text;
	wchar_t *p_wchar = char_to_wchar(str);

	::CoInitialize(NULL); // COM��ʼ��
	CLSID CLSID_SpVoice;
	CLSIDFromProgID(L"SAPI.SpVoice", &CLSID_SpVoice);
	ISpVoice *pSpVoice = NULL;
	IEnumSpObjectTokens *pSpEnumTokens = NULL;

	// ��ȡISpVoice�ӿ�
	if (FAILED(CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_INPROC_SERVER, IID_ISpVoice, (void**)&pSpVoice)))
	{
		cout << "error:��ȡISpVoice�ӿ�ʧ��" << endl;
		return -1;
	}
	/* �����ʶ����� */
	pSpVoice->SetVolume(voice_opts.volume);
	/* �����ʶ��ٶ� */
	pSpVoice->SetRate(voice_opts.rate);
	/* ����ͬ���ʶ���ʱ�¼�����λ��ms */
	pSpVoice->SetSyncSpeakTimeout(5000);

	// �о����е�����token������ͨ��pSpEnumTokensָ��Ľӿڵõ�
	if (SUCCEEDED(SpEnumTokens(SPCAT_VOICES, NULL, NULL, &pSpEnumTokens)))
	{
		ULONG count = 0;
		pSpEnumTokens->GetCount(&count);
		// �жϱ�������token�����Ƿ�������1��
		if (count >= 1)
		{
			ISpObjectToken *pSpToken = NULL;
			pSpEnumTokens->Item(0, &pSpToken);
			pSpVoice->SetVoice(pSpToken); // ���õ�ǰ����tokenΪpSpToken
			for (int i = 0; i < voice_opts.forTime; i++)
				pSpVoice->Speak((LPCWSTR)p_wchar, SPF_DEFAULT, NULL); // �ʶ����ĺ�Ӣ�ĵĻ���ַ���
			pSpToken->Release(); // �ͷ�token
		}

		//// ���λ�ȡÿ��token���ʶ��ַ���
		//while (SUCCEEDED(pSpEnumTokens->Next(1, &pSpToken, NULL)) && pSpToken != NULL)
		//{
		//	pSpVoice->SetVoice(pSpToken);      // ���õ�ǰ����tokenΪpSpToken
		//	pSpVoice->Speak(L"Hello Word �������", SPF_DEFAULT, NULL);     // �ʶ����ĺ�Ӣ�ĵĻ���ַ���
		//	pSpToken->Release();       // �ͷ�token
		//}

		pSpEnumTokens->Release();        // �ͷ�pSpEnumTokens�ӿ�
	}

	delete[] p_wchar;
	pSpVoice->Release();
	::CoUninitialize();

	return 0;
}