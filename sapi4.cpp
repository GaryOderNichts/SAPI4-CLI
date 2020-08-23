#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <string.h>
#include <stdio.h>
#include <mmsystem.h>
#include <initguid.h>
#include <objbase.h>
#include <objerror.h>
#include <ole2ver.h>

#include <winsock2.h>
#include <ws2tcpip.h>
#include <math.h>
#include <time.h>

#include <speech.h>

#include "argparse.h"

class CTestNotify : public ITTSNotifySink {
	public:
		CTestNotify (void);
		~CTestNotify (void);

		// IUnkown members that delegate to m_punkOuter
		// Non-delegating object IUnknown
		STDMETHODIMP		   QueryInterface (REFIID, LPVOID FAR *);
		STDMETHODIMP_(ULONG) AddRef(void);
		STDMETHODIMP_(ULONG) Release(void);

		// ITTSNotifySink
		STDMETHOD (AttribChanged)  (DWORD);
		STDMETHOD (AudioStart)	   (QWORD);
		STDMETHOD (AudioStop)	   (QWORD);
		STDMETHOD (Visual)		   (QWORD, CHAR, CHAR, DWORD, PTTSMOUTH);
};
typedef CTestNotify *PCTestNotify;


HINSTANCE		ghInstance;
PITTSCENTRAL	gpITTSCentral = NULL;
PITTSATTRIBUTES gpITTSAttributes = NULL;
CTestNotify		gNotify;
WORD			gwRealTime = 0xFFFFFFFF;
PIAUDIOFILE		gpIAF = NULL;
BOOL			SAMDone = FALSE;


CTestNotify::CTestNotify()
{
}

CTestNotify::~CTestNotify()
{
// this space intentionally left blank
}

STDMETHODIMP CTestNotify::QueryInterface(
	REFIID riid,
	LPVOID *ppv)
{
	*ppv = NULL;

	/* always return our IUnknown for IID_IUnknown */
	if (IsEqualIID (riid, IID_IUnknown) || IsEqualIID(riid,IID_ITTSNotifySink)) {
		*ppv = (LPVOID) this;
		return S_OK;
	}

	// otherwise, cant find
	return ResultFromScode(E_NOINTERFACE);
}

STDMETHODIMP_ (ULONG) CTestNotify::AddRef()
{
	return 1;
}

STDMETHODIMP_(ULONG) CTestNotify::Release()
{
	return 1;
}

STDMETHODIMP CTestNotify::AttribChanged(
	DWORD dwAttribID)
{
	return NOERROR;
}

STDMETHODIMP CTestNotify::AudioStart(
	QWORD qTimeStamp)
{
	return NOERROR;
}

STDMETHODIMP CTestNotify::AudioStop(
	QWORD qTimeStamp)
{
	gpIAF->Flush();
	SAMDone = TRUE;

	return NOERROR;
}

STDMETHODIMP CTestNotify::Visual(
	QWORD qTimeStamp,
	CHAR cIPAPhoneme,
	CHAR cEnginePhoneme,
	DWORD dwHints,
	PTTSMOUTH pTTSMouth)
{
	return NOERROR;
}

PITTSCENTRAL FindAndSelect(PTTSMODEINFO pTTSInfo)
{
	HRESULT hRes;
	TTSMODEINFO ttsResult;
	PITTSFIND pITTSFind;
	PITTSCENTRAL pITTSCentral;

	hRes = CoCreateInstance(CLSID_TTSEnumerator, NULL, CLSCTX_ALL, IID_ITTSFind, (void**)&pITTSFind);
	if (FAILED(hRes))
		return NULL;

	hRes = pITTSFind->Find(pTTSInfo, NULL, &ttsResult);
	if (FAILED(hRes))
		goto failRelease;

	if (strcmp(ttsResult.szModeName, pTTSInfo->szModeName) != 0) 
	{
		printf("Error: Voice \"%s\" not found\n", pTTSInfo->szModeName);
		printf("Available voices are:\n");

		PITTSENUM pITTSEnum;
		hRes = CoCreateInstance(CLSID_TTSEnumerator, NULL, CLSCTX_ALL, IID_ITTSEnum, (void**)&pITTSEnum);
		if (FAILED(hRes))
			goto failRelease;
	  
		TTSMODEINFO nTTSInfo;
		while (!pITTSEnum->Next(1, &nTTSInfo, NULL)) 
		{
			printf("%s\n", nTTSInfo.szModeName);
		}
	
		pITTSEnum->Release();

		return NULL;
	}
	
	hRes = CoCreateInstance(CLSID_AudioDestFile, NULL, CLSCTX_ALL, IID_IAudioFile, (void**)&gpIAF);
	if (FAILED(hRes))
		goto failRelease;

	hRes = pITTSFind->Select(ttsResult.gModeID, &pITTSCentral, gpIAF);
	if (FAILED(hRes))
		goto failRelease;

	printf("Selected voice \"%s\"\n", ttsResult.szModeName);
	
	pITTSFind->Release();

	return pITTSCentral;
	
failRelease:
	pITTSFind->Release();
	return NULL;
}

BOOL BeginOLE()
{
	DWORD dwVer = CoBuildVersion();

	if (rmm != HIWORD(dwVer))
		return FALSE;

	if (FAILED(CoInitialize(NULL)))
		return FALSE;

	return TRUE;
}

BOOL EndOLE()
{
	CoUninitialize();
	return TRUE;
}

int main(int argc, char** argv)
{
	argparse::ArgumentParser parser("SAPI4", "Use SAPI4 without a webinterface");
	parser.add_argument()
		.names({"-v", "--voice"})
		.description("voice")
		.required(true);
	parser.add_argument()
		.names({"-s", "--speed"})
		.description("speed")
		.required(true);
	parser.add_argument()
		.names({"-p", "--pitch"})
		.description("pitch")
		.required(true);
	parser.add_argument()
		.names({"-t", "--text"})
		.description("text")
		.required(true);
	parser.add_argument()
		.names({"-o", "--output"})
		.description("output file")
		.required(true);
	parser.enable_help();
	
	auto argerr = parser.parse(argc, (const char**) argv);
	if (argerr)
	{
		std::cout << argerr << std::endl;
		return -1;
	}

	if (parser.exists("help")) 
	{
		parser.print_help();
		return 0;
	}

	srand(time(NULL));
	if (!BeginOLE())
		return 1;

	TTSMODEINFO	ModeInfo;
	SDATA data;

	// find the right object
	memset(&ModeInfo, 0, sizeof(ModeInfo));
	strcpy(ModeInfo.szModeName, parser.get<std::string>("v").c_str());
	
	gpITTSCentral = FindAndSelect(&ModeInfo);
	if (!gpITTSCentral) 
	{
		std::cout << "Error selecting voice" << std::endl;
		return -1;
	}
	
	gpIAF->RealTimeSet(gwRealTime);

	if (FAILED(gpITTSCentral->QueryInterface(IID_ITTSAttributes, (void**)&gpITTSAttributes))) 
	{
		std::cout << "Query interface failed" << std::endl;
		return -1;
	}
	
	WORD minPitch, maxPitch;
	DWORD minSpeed, maxSpeed;
	gpITTSAttributes->PitchSet(TTSATTR_MINPITCH);
	gpITTSAttributes->SpeedSet(TTSATTR_MINSPEED);
	gpITTSAttributes->PitchGet(&minPitch);
	gpITTSAttributes->SpeedGet(&minSpeed);
	gpITTSAttributes->PitchSet(TTSATTR_MAXPITCH);
	gpITTSAttributes->SpeedSet(TTSATTR_MAXSPEED);
	gpITTSAttributes->PitchGet(&maxPitch);
	gpITTSAttributes->SpeedGet(&maxSpeed);

	int pitch = parser.get<int>("p");
	int speed = parser.get<int>("s");

	if (speed < minSpeed)
	{
		std::cout << "Warning: Speed " << speed << " is below the minimum speed " << minSpeed << std::endl;
		std::cout << "\t Using minimum speed" << std::endl;
		speed = minSpeed;
	}
	else if (speed > maxSpeed)
	{
		std::cout << "Warning: Speed " << speed << " is above the maximum speed " << maxSpeed << std::endl;
		std::cout << "\t Using maximum speed" << std::endl;
		speed = maxSpeed;
	}
	
	if (pitch < minPitch)
	{
		std::cout << "Warning: Pitch " << pitch << " is below the minimum pitch " << minPitch << std::endl;
		std::cout << "\t Using minimum pitch" << std::endl;
		pitch = minPitch;
	}
	else if (pitch > maxPitch)
	{
		std::cout << "Warning: Pitch " << pitch << " is above the maximum pitch " << maxPitch << std::endl;
		std::cout << "\t Using maximum pitch" << std::endl;
		pitch = maxPitch;
	}

	gpITTSAttributes->PitchSet(pitch);
	gpITTSAttributes->SpeedSet(speed);

	std::cout << "Pitch: " << pitch << " Speed: " << speed << std::endl;
	
	DWORD dwRegKey;
	gpITTSCentral->Register((void*)&gNotify, IID_ITTSNotifySink, &dwRegKey);

	WCHAR wszFile[17] = L"";
	MultiByteToWideChar(CP_ACP, 0, parser.get<std::string>("o").c_str(), -1, wszFile, sizeof(wszFile) / sizeof(WCHAR));
	if (gpIAF->Set(wszFile, 1)) 
	{
		std::cout << "Setting output file failed" << std::endl;
		return -1;
	}
	
	std::string text = parser.get<std::string>("t");
	data.dwSize = text.size();
	data.pData = (void*) text.c_str();
	gpITTSCentral->AudioReset();
	gpITTSCentral->TextData(CHARSET_TEXT, TTSDATAFLAG_TAGGED, data, NULL, IID_ITTSBufNotifySink);

	BOOL fGotMessage;
	MSG msg;
	while ((fGotMessage = GetMessage(&msg, (HWND) NULL, 0, 0)) != 0 && fGotMessage != -1 && !SAMDone) 
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	gpIAF->Flush();
	SAMDone = FALSE;
	
	std::cout << "Done!" << std::endl;

	WSACleanup();
	
	gpIAF->Release();
	gpITTSAttributes->Release();
	gpITTSCentral->Release();

	if (!EndOLE()) 
	{
		printf("EndOLE failed\n");
		return 1;
	}
	
	return 0;
}
