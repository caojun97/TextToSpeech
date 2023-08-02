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
	char text[512]; // 朗读文本
	int  volume;    // 朗读音量，有效范围[0,100]
	int  rate;      // 朗读速率，有效范围[-10,10]
	int  forTime;   // 朗读循环次数

}VOICE_OPTS;

static void lib_tts_usage(const char *program)
{
	cout << endl;
	cout << program << " [-v volume] [-r rate] [-f forTime] {-t text}\n" << endl;
	cout << "描述：使用该工具可以将文本转换成语音" << endl;
	cout << "参数列表：\n"
			"  -v --volume   朗读音量(int)，有效范围[0,100]\n" 
			"  -r --rate     朗读速率(int)，有效范围[-10,10]\n"
			"  -t --text     朗读文本(string)，支持中英混读\n"
			"  -? --help     帮助文档\n"
			"  -f --forTime  朗读循环次数(int)"<< endl;
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
	/* VOICE_OPTS 变量初始化 */
	lib_tts_opt_init(&voice_opts);
	/* 命令行参数读取 */
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

	::CoInitialize(NULL); // COM初始化
	CLSID CLSID_SpVoice;
	CLSIDFromProgID(L"SAPI.SpVoice", &CLSID_SpVoice);
	ISpVoice *pSpVoice = NULL;
	IEnumSpObjectTokens *pSpEnumTokens = NULL;

	// 获取ISpVoice接口
	if (FAILED(CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_INPROC_SERVER, IID_ISpVoice, (void**)&pSpVoice)))
	{
		cout << "error:获取ISpVoice接口失败" << endl;
		return -1;
	}
	/* 调节朗读音量 */
	pSpVoice->SetVolume(voice_opts.volume);
	/* 调节朗读速度 */
	pSpVoice->SetRate(voice_opts.rate);
	/* 设置同步朗读超时事件，单位：ms */
	pSpVoice->SetSyncSpeakTimeout(5000);

	// 列举所有的语音token，可以通过pSpEnumTokens指向的接口得到
	if (SUCCEEDED(SpEnumTokens(SPCAT_VOICES, NULL, NULL, &pSpEnumTokens)))
	{
		ULONG count = 0;
		pSpEnumTokens->GetCount(&count);
		// 判断本地语音token数量是否至少有1个
		if (count >= 1)
		{
			ISpObjectToken *pSpToken = NULL;
			pSpEnumTokens->Item(0, &pSpToken);
			pSpVoice->SetVoice(pSpToken); // 设置当前语音token为pSpToken
			for (int i = 0; i < voice_opts.forTime; i++)
				pSpVoice->Speak((LPCWSTR)p_wchar, SPF_DEFAULT, NULL); // 朗读中文和英文的混合字符串
			pSpToken->Release(); // 释放token
		}

		//// 依次获取每个token并朗读字符串
		//while (SUCCEEDED(pSpEnumTokens->Next(1, &pSpToken, NULL)) && pSpToken != NULL)
		//{
		//	pSpVoice->SetVoice(pSpToken);      // 设置当前语音token为pSpToken
		//	pSpVoice->Speak(L"Hello Word 世界你好", SPF_DEFAULT, NULL);     // 朗读中文和英文的混合字符串
		//	pSpToken->Release();       // 释放token
		//}

		pSpEnumTokens->Release();        // 释放pSpEnumTokens接口
	}

	delete[] p_wchar;
	pSpVoice->Release();
	::CoUninitialize();

	return 0;
}