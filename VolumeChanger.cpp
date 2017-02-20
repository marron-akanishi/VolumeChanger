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

	// �f�o�C�X�񋓃I�u�W�F�N�g���擾����
	ComPtr<IMMDeviceEnumerator> deviceEnumerator;
	HRESULT hr = CoCreateInstance(
		__uuidof(MMDeviceEnumerator),
		nullptr,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&deviceEnumerator));
	if (FAILED(initializer))
		RaiseException(hr);

	// ����̃}���`���f�B�A�o�̓f�o�C�X�i�X�s�[�J�[�j���擾����
	ComPtr<IMMDevice> device;
	hr = deviceEnumerator->GetDefaultAudioEndpoint(
		EDataFlow::eRender,
		ERole::eMultimedia,
		&device);
	if (FAILED(initializer))
		RaiseException(hr);

	// �I�[�f�B�I�G���h�|�C���g�̃{�����[���I�u�W�F�N�g���쐬����
	ComPtr<IAudioEndpointVolume> audioEndpointVolume;
	hr = device->Activate(
		__uuidof(IAudioEndpointVolume),
		CLSCTX_INPROC_SERVER,
		nullptr,
		&audioEndpointVolume);
	if (FAILED(hr))
		RaiseException(hr);

	// �W���b�N�܂ł̏����擾����
	ComPtr<IDeviceTopology> pDeviceTopology;
	ComPtr<IConnector> pConnEP;
	IConnector* pConnDeviceTo;
	ComPtr<IPart> pPart;
	device->Activate(__uuidof(IDeviceTopology), CLSCTX_INPROC_SERVER, NULL, (void**)&pDeviceTopology);
	pDeviceTopology->GetConnector(0, &pConnEP);
	pConnEP->GetConnectedTo(&pConnDeviceTo);
	pConnDeviceTo->QueryInterface(__uuidof(IPart), (void**)&pPart);

	// ���ʒ���������
	ComPtr<IKsJackDescription> pJackDesc = NULL;
	UINT nNumJacks = 0;
	UINT SpeakerCount = 1;
	float HeadphoneVol = 0.1;
	float SpeakerVol = 0.2;
	BOOL SpeakerMute = TRUE;
	UINT LoopDelay = 200;
	UINT oldNumJacks = 0;
	// �ݒ�̓ǂݍ���
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
	// ���C�����[�v
	while (1) {
		pPart->Activate(CLSCTX_INPROC_SERVER, __uuidof(IKsJackDescription), &pJackDesc);
		pJackDesc->GetJackCount(&nNumJacks);
		if (nNumJacks != oldNumJacks) {
			// �N�����̏���
			if (oldNumJacks == 0) {
				if (nNumJacks == SpeakerCount) {
					hr = audioEndpointVolume->GetMute(&SpeakerMute);
					hr = audioEndpointVolume->GetMasterVolumeLevelScalar(&SpeakerVol); //�X�s�[�J�[���ʂ̎擾
				} else {
					hr = audioEndpointVolume->GetMasterVolumeLevelScalar(&HeadphoneVol); //�w�b�h�z�����ʂ̎擾
				}
				MessageBox(NULL, TEXT("�N�����܂���\n�f�o�C�X�̔����������\�ł�"), TEXT("VolumeChanger"), MB_OK);
			} else {
				if (nNumJacks == SpeakerCount) {
					hr = audioEndpointVolume->GetMasterVolumeLevelScalar(&HeadphoneVol); //�w�b�h�z�����ʂ̎擾
					hr = audioEndpointVolume->SetMute(SpeakerMute, nullptr); //�~���[�g��Ԃ�ݒ�
					hr = audioEndpointVolume->SetMasterVolumeLevelScalar(SpeakerVol, nullptr);
				} else {
					hr = audioEndpointVolume->GetMute(&SpeakerMute); //�X�s�[�J�[�ł̃~���[�g��Ԃ��擾
					hr = audioEndpointVolume->GetMasterVolumeLevelScalar(&SpeakerVol); //�X�s�[�J�[���ʂ̎擾
					hr = audioEndpointVolume->SetMute(FALSE, nullptr); //�~���[�g����
					hr = audioEndpointVolume->SetMasterVolumeLevelScalar(HeadphoneVol, nullptr);
				}
			}
			oldNumJacks = nNumJacks;
		}
		Sleep(LoopDelay);
	}

	return 0;
}