#include "stdafx.h"

#define GetBool(buf) strcmp(buf,"True") ? TRUE : FALSE

using namespace::Microsoft::WRL;
using namespace::Microsoft::WRL::Details;

class ComInitializer {
private:
	HRESULT m_hr;
public:
	ComInitializer() { m_hr = CoInitialize(nullptr); }
	ComInitializer(LPVOID pvReserved) { m_hr = CoInitialize(pvReserved); }
	~ComInitializer() { CoUninitialize(); }
	operator HRESULT() const { return m_hr; }
};

int WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int) {
	ComInitializer initializer;
	if (FAILED(initializer))
		return -1;

	// デバイス列挙オブジェクトを取得する
	ComPtr<IMMDeviceEnumerator> deviceEnumerator;
	HRESULT hr = CoCreateInstance(
		__uuidof(MMDeviceEnumerator),
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&deviceEnumerator));
	if (FAILED(initializer))
		RaiseException(hr);

	// 既定のマルチメディア出力デバイス（スピーカー）を取得する
	ComPtr<IMMDevice> device;
	hr = deviceEnumerator->GetDefaultAudioEndpoint(
		EDataFlow::eRender,
		ERole::eMultimedia,
		&device);
	if (FAILED(initializer))
		RaiseException(hr);

	// オーディオエンドポイントのボリュームオブジェクトを作成する
	ComPtr<IAudioEndpointVolume> audioEndpointVolume;
	hr = device->Activate(
		__uuidof(IAudioEndpointVolume),
		CLSCTX_INPROC_SERVER,
		nullptr,
		&audioEndpointVolume);
	if (FAILED(hr))
		RaiseException(hr);

	// ジャックまでの情報を取得する
	ComPtr<IDeviceTopology> pDeviceTopology;
	ComPtr<IConnector> pConnEP;
	IConnector* pConnDeviceTo;
	ComPtr<IPart> pPart;
	device->Activate(__uuidof(IDeviceTopology), CLSCTX_INPROC_SERVER, NULL, (void**)&pDeviceTopology);
	pDeviceTopology->GetConnector(0, &pConnEP);
	pConnEP->GetConnectedTo(&pConnDeviceTo);
	pConnDeviceTo->QueryInterface(__uuidof(IPart), (void**)&pPart);

	// 音量調整をする
	ComPtr<IKsJackDescription> pJackDesc = NULL;
	UINT nNumJacks = 0;
	UINT SpeakerCount = 1;
	float HeadphoneVol = 0.1;
	float SpeakerVol = 0.2;
	BOOL SpeakerMute = TRUE;
	UINT LoopDelay = 200;
	UINT oldNumJacks = 0;
	// 設定の読み込み
	char buf[64];
	GetPrivateProfileString(TEXT("Default"), TEXT("SpeakerCount"), TEXT("1"), buf, 64, TEXT("./setting.ini"));
	SpeakerCount = atoi(buf);
	GetPrivateProfileString(TEXT("Default"), TEXT("SpeakerVol"), TEXT("0.2"), buf, 64, TEXT("./setting.ini"));
	SpeakerVol = atof(buf);
	GetPrivateProfileString(TEXT("Default"), TEXT("SpeakerMute"), TEXT("True"), buf, 64, TEXT("./setting.ini"));
	SpeakerMute = GetBool(buf);
	GetPrivateProfileString(TEXT("Default"), TEXT("HeadphoneVol"), TEXT("0.1"), buf, 64, TEXT("./setting.ini"));
	HeadphoneVol = atof(buf);
	GetPrivateProfileString(TEXT("Default"), TEXT("LoopDelay"), TEXT("200"), buf, 64, TEXT("./setting.ini"));
	LoopDelay = atoi(buf);
	// メインループ
	while (1) {
		pPart->Activate(CLSCTX_INPROC_SERVER, __uuidof(IKsJackDescription), &pJackDesc);
		pJackDesc->GetJackCount(&nNumJacks);
		if (nNumJacks != oldNumJacks) {
			// 起動時の処理
			if (oldNumJacks == 0) {
				if (nNumJacks == SpeakerCount) {
					hr = audioEndpointVolume->GetMute(&SpeakerMute);
					hr = audioEndpointVolume->GetMasterVolumeLevelScalar(&SpeakerVol); //スピーカー音量の取得
				} else {
					hr = audioEndpointVolume->GetMasterVolumeLevelScalar(&HeadphoneVol); //ヘッドホン音量の取得
				}
				MessageBox(NULL, TEXT("起動しました\nデバイスの抜き差しが可能です"), TEXT("VolumeChanger"), MB_OK);
			} else {
				if (nNumJacks == SpeakerCount) {
					hr = audioEndpointVolume->GetMasterVolumeLevelScalar(&HeadphoneVol); //ヘッドホン音量の取得
					hr = audioEndpointVolume->SetMute(SpeakerMute, nullptr); //ミュート状態を設定
					hr = audioEndpointVolume->SetMasterVolumeLevelScalar(SpeakerVol, nullptr);
				} else {
					hr = audioEndpointVolume->GetMute(&SpeakerMute); //スピーカーでのミュート状態を取得
					hr = audioEndpointVolume->GetMasterVolumeLevelScalar(&SpeakerVol); //スピーカー音量の取得
					hr = audioEndpointVolume->SetMute(FALSE, nullptr); //ミュート解除
					hr = audioEndpointVolume->SetMasterVolumeLevelScalar(HeadphoneVol, nullptr);
				}
			}
			oldNumJacks = nNumJacks;
		}
		Sleep(LoopDelay);
	}

	return 0;
}