#include "stdafx.h"

#define USE_SPEECH_DLL
#define LANE_SPEECH_EXPORTS

#include "LaneSpeech.h"

#ifdef USE_SPEECH_DLL	//������USE_SPEECH_DLL���Ͱ�����DLL����������������
BOOL APIENTRY DllMain( HANDLE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved )
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:  
			CoInitialize(NULL);  
			break;
		case DLL_THREAD_ATTACH:   
			break;
		case DLL_THREAD_DETACH:   
			break;
		case DLL_PROCESS_DETACH:  
			CoUninitialize();    
			break;
    }
    return TRUE;
}
#endif	//USE_SPEECH_DLL


CSR::CSR()
{
	m_bRecoSuccess	= false;
	m_bRecoEnd		= true;
	m_pSREngine		= NULL;
	m_pSRContext	= NULL;
	m_pSRGrammar	= NULL;	
	m_pInputStream	= NULL;
	m_pToken		= NULL;	
	m_pAudio		= NULL;
	UpdateGrammerID ();
}

CSR::~CSR()
{
	if ( m_pToken ) { 
		m_pToken.Release();
		m_pToken = NULL;
	}
	if ( m_pAudio ) { 
		m_pAudio.Release();
		m_pAudio = NULL;
	}
	if ( m_pInputStream ) { 
		m_pInputStream.Release();
		m_pInputStream = NULL;
	}
	if ( m_pSRGrammar ) { 
		m_pSRGrammar.Release();
		m_pSRGrammar = NULL;
	}
	if ( m_pSRContext ) { 
		m_pSRContext.Release();
		m_pSRContext = NULL;
	}
	if ( m_pSREngine ) { 
		m_pSREngine.Release();
		m_pSREngine = NULL;
	}
}

// ����һ��SR���棬 dwLanguageΪ�������࣬bInprocΪ ��ʶ�������Ƿ�Ϊ���������
bool CSR::Create ( const DWORD dwLanguage,
				   const bool bInproc )
{
	HRESULT hr;

	if ( bInproc ) {//���������
		hr = m_pSREngine.CoCreateInstance ( CLSID_SpInprocRecognizer );
		if ( !::CheckHr(hr, "CoCreateInstance()", "Create()") ) 
			return false; 
		
		hr = SpCreateDefaultObjectFromCategoryId ( SPCAT_AUDIOIN, &m_pAudio );	//����һ��Ĭ����Ƶ��
		if ( !::CheckHr(hr, "SpCreateDefaultObjectFromCategoryId()", "Create()") ) 
			return false; 
		
		hr = m_pSREngine->SetInput ( m_pAudio, TRUE );	//Ϊʶ������������Ƶ����
		if ( !::CheckHr(hr, "SetInput()", "Create()") ) 
			return false; 
	}
	else {//���������
		hr = m_pSREngine.CoCreateInstance ( CLSID_SpSharedRecognizer );//����һ�������SR���档
		if ( !::CheckHr(hr, "CoCreateInstance", "Create()") ) 
			return false; 
	}
	
	if ( dwLanguage == SP_CHINESE )  //ʶ������Ϊ����
		hr = SpFindBestToken ( SPCAT_RECOGNIZERS, L"language=804", 
								NULL, &m_pToken );
	else	//ʶ������Ϊ����
		hr = SpFindBestToken ( SPCAT_RECOGNIZERS, L"language=409", 
								NULL, &m_pToken );
	if ( !::CheckHr(hr, "SpFindBestToken()", "Create()") )
		return false;
	
	hr = m_pSREngine->SetRecognizer ( m_pToken );	//�����������
	if ( !::CheckHr(hr, "SetRecognizer()", "Create()") )
		return false;

	hr = m_pSREngine->CreateRecoContext ( &m_pSRContext );	//����������
	if ( !::CheckHr(hr, "CreateRecoContext()", "Create()") )
		return false;

	hr = m_pSRContext->SetNotifyWin32Event();			//����win32�¼�
	if ( !::CheckHr(hr, "SetNotifyWin32Event()", "Create()") )
		return false;

	//����Ҫ������¼�
	hr = m_pSRContext->SetInterest ( SPFEI(SPEI_RECOGNITION) | SPFEI( SPEI_FALSE_RECOGNITION ) |
									SPFEI(SPEI_END_SR_STREAM) | SPFEI(SPEI_START_SR_STREAM) |
									SPFEI(SPEI_SOUND_END) | SPFEI( SPEI_SOUND_START ),
									SPFEI(SPEI_RECOGNITION) | SPFEI( SPEI_FALSE_RECOGNITION ) |
									SPFEI(SPEI_END_SR_STREAM) | SPFEI(SPEI_START_SR_STREAM) | 
									SPFEI(SPEI_SOUND_END) | SPFEI( SPEI_SOUND_START ));
	if ( !::CheckHr(hr, "SetInterest()", "Create()") )
		return false;
	
	return true;
}

// ����ʶ������Ҫ�õĵ��ķ��ļ�
bool CSR::LoadCrammarFile ( const WCHAR *pwchFilename )
{
	HRESULT hr;
	
	hr = m_pSRContext->CreateGrammar ( ullGrammerID, &m_pSRGrammar );	//�����ķ��ӿ�
	if ( !::CheckHr(hr, "CreateGrammar()", "LoadCrammarFile()") )
		return false;
	
	hr = m_pSRGrammar->LoadCmdFromFile ( pwchFilename, SPLO_DYNAMIC );	//�����ķ��ļ�
	if ( !::CheckHr(hr, "LoadCmdFromFile()", "LoadCrammarFile()") )
		return false;

	return true;
}

// ��ʶ����������Ϊ��wav�ļ�������룬pwchFileNameΪ�ļ�����FileFormatΪ�ļ���ʽ
bool CSR::SetInputFromWav ( const WCHAR *pwchFileName, const SPSTREAMFORMAT FileFormat )
{
	HRESULT hr;

	hr = m_pSRGrammar->SetRuleState ( NULL, NULL, SPRS_INACTIVE );	//ȡ�����򼤻�
	if ( !::CheckHr( hr, "SetRuleState()", "SetInputFromWav()" ) )  
		return false;
	
	if ( m_pInputStream ) {
		m_pInputStream.Release();
		m_pInputStream = NULL;
	}
	
	hr = m_pInputStream.CoCreateInstance(CLSID_SpStream);			//�������ӿ�
	if ( !::CheckHr( hr, "CoCreateInstance()", "SetInputFromWav()" ) )  
		return false;
	
	CSpStreamFormat sInputFormat;
	//hr = sInputFormat.AssignFormat(SPSF_22kHz16BitStereo);			//����wav�ļ���ʽ
	hr = sInputFormat.AssignFormat ( FileFormat );
	if ( !::CheckHr( hr, "AssignFormat()", "SetInputFromWav()" ) ) 
		return false;
	
	hr = m_pInputStream->BindToFile	(	pwchFileName,				//��wav�ļ�����
										SPFM_OPEN_READONLY,
										&sInputFormat.FormatId(),
										sInputFormat.WaveFormatExPtr(),
										SPFEI_ALL_EVENTS  );
	if ( !::CheckHr( hr, "BindToFile()", "SetInputFromWav()" ) ) 	
		return false;
	
	hr = m_pSREngine->SetInput ( m_pInputStream, true );			//�������������
	if ( !::CheckHr( hr, "SetInput()", "SetInputFromWav()" ) ) 
		return false;
	
	//����Ҫ������¼�
	hr = m_pSRContext->SetInterest (SPFEI(SPEI_RECOGNITION) | SPFEI( SPEI_FALSE_RECOGNITION ) |
									SPFEI(SPEI_END_SR_STREAM) | SPFEI(SPEI_START_SR_STREAM) |
									SPFEI(SPEI_SOUND_END) | SPFEI( SPEI_SOUND_START ),
									SPFEI(SPEI_RECOGNITION) | SPFEI( SPEI_FALSE_RECOGNITION ) |
									SPFEI(SPEI_END_SR_STREAM) | SPFEI(SPEI_START_SR_STREAM) |
									SPFEI(SPEI_SOUND_END) | SPFEI( SPEI_SOUND_START ));
	
	return true;
}

// ʶ�����Ĵ�����̣�ʶ������ʶ�����Ĺ��̾ͷ����������Ϊ��ʱֵ
void CSR::ProcessReco( const DWORD dwMilliseconds )
{
	//�����¼���������������timeoutӦ�����õı�audio������
	//������������һ�������ʱ�������ڲ���ʶ����ǰ��
	bool fEndStreamReached = false;
	bool bTimeOut = false;
	while (!fEndStreamReached && S_OK == m_pSRContext->WaitForNotifyEvent(dwMilliseconds))
	{
		CSpEvent spEvent;
		while (!fEndStreamReached && S_OK == spEvent.GetFrom(m_pSRContext))
		{
			switch (spEvent.eEventId)
			{
				case SPEI_RECOGNITION:
					//ShowMsg("ʶ��ɹ�");
					m_bRecoSuccess = true;
					OnRecoSuccess ( spEvent.RecoResult() );
					//m_bRecoEnd	   = true;
					break;
				case SPEI_FALSE_RECOGNITION:
					//ShowMsg("ʶ��ʧ��");
					m_bRecoSuccess = false;
					//m_bRecoEnd	   = true;
					break;
				
				case SPEI_END_SR_STREAM:
					m_bRecoEnd			= true;
					fEndStreamReached	= true;
					//ShowMsg("ʶ�����");
					bTimeOut = false;
					break;

				case SPEI_START_SR_STREAM:
					//ShowMsg("ʶ��ʼ");
					m_bRecoEnd	   = false;
					bTimeOut = true;
					break;

				//case SPEI_SOUND_START:
				//	m_bRecoEnd			= true;
					//fEndStreamReached	= true;
				//	ShowMsg("������ʼ");
				//	break;

				//case SPEI_SOUND_END:
				//	m_bRecoEnd			= true;
					//fEndStreamReached	= true;
				//	ShowMsg("��������");
				//	break;
			}
			spEvent.Clear();
		}
	}
	if ( bTimeOut || !fEndStreamReached ) {
		m_bRecoEnd	= true;
	}

}

// ʶ��ɹ�������£��ͱ���ʶ��Ķ���
void CSR::OnRecoSuccess ( ISpPhrase *ppPhrase )
{
	SPPHRASE* pPhrase = NULL;

	if ( SUCCEEDED ( ppPhrase->GetPhrase (&pPhrase) ) ) {
		PWCHAR pwszTemp = NULL;
		if ( pPhrase->Rule.ulId == 3000 ) {
			if ( pPhrase->pProperties->ulId == 3001 ) {
				HRESULT hr = ppPhrase->GetText ( pPhrase->pProperties->ulFirstElement,
											pPhrase->pProperties->ulCountOfElements,
											FALSE,
											&pwszTemp,
											NULL );
				
				if ( hr != S_OK ) {
					m_bRecoSuccess = false;
				}
				
				wcscpy ( m_wchRecoWord, pwszTemp );
			}
		}
		if ( pPhrase ) {  // �ͷŶ������
			::CoTaskMemFree(pPhrase);
		}
	}
}

// ����ʶ���Ƿ����
bool CSR::IsRecoEnd() const
{
	return	m_bRecoEnd;
}

// ����ʶ���Ƿ�ɹ�
bool CSR::IsRecoSuccess() const
{
	return	m_bRecoSuccess;
}

// ��ʼʶ��
bool CSR::StartReco()
{
	HRESULT hr;
	hr = m_pSRGrammar->SetRuleState ( NULL, NULL, SPRS_ACTIVE );	//�������й���
	if ( !::CheckHr( hr, "SetRuleState()", "SetInputFromWav()" ) )  
		return false;
	return true;
}

// ���ʶ��Ķ�� UNICODE �ַ��汾
void CSR::GetRecoWordW ( wchar_t * pwchRecoWord )
{
	wcscpy ( pwchRecoWord, m_wchRecoWord );
}

// ���ʶ��Ķ�� ANSI �ַ��汾
void CSR::GetRecoWordA (char * pchRecoWord )
{
	char temp[128];
	WideCharToMultiByte ( CP_ACP, //��charת����wchar_t
						  0,
						  m_wchRecoWord,
						  sizeof(m_wchRecoWord) / sizeof(wchar_t),
						  temp,
						  128 * sizeof( char ),
						  NULL,
						  NULL );
	strcpy ( pchRecoWord, temp );
}

ULONGLONG	CSR::ullGrammerID = 1000;
	
/////////////////////////////////////////////////////////////
//
//	CTTS ��
//
/////////////////////////////////////////////////////////////
CTTS::CTTS()
{
	m_pVoice		= NULL;
	m_pToken		= NULL;
	m_pOutputStream	= NULL;
	m_pAudio		= NULL;
}

CTTS::~CTTS()
{
	if ( m_pToken ) { 
		m_pToken.Release();
		m_pToken = NULL;
	}
	if ( m_pAudio ) { 
		m_pAudio.Release();
		m_pAudio = NULL;
	}
	if ( m_pOutputStream ) { 
		m_pOutputStream.Release();
		m_pOutputStream = NULL;
	}
	if ( m_pVoice ) { 
		m_pVoice.Release();
		m_pVoice = NULL;
	}
}

// ����һ��CTTS����
bool CTTS::Create( const DWORD dwLanguage )
{
	HRESULT hr;
	if ( m_pVoice )
		return true;
	hr = m_pVoice.CoCreateInstance ( CLSID_SpVoice );
	if ( !::CheckHr( hr, "CoCreateInstance()", "Create" ) )  
		return false;
	
	SetLanguage ( dwLanguage );
	
	hr = SpCreateDefaultObjectFromCategoryId ( SPCAT_AUDIOIN, &m_pAudio );//����һ��Ĭ����Ƶ��
	if ( !::CheckHr ( hr, "CreateDefaultObjectFodd()", "Create" ) ) 
		return false;
	
	hr = m_pVoice->SetInterest( SPFEI( SPEI_START_INPUT_STREAM ) |
								SPFEI( SPEI_END_INPUT_STREAM ),
    							SPFEI( SPEI_START_INPUT_STREAM ) |
    							SPFEI( SPEI_END_INPUT_STREAM ) );
	if ( !::CheckHr ( hr, "SetInterest()", "Create" ) )  
		return false;
	
	//hr = m_pVoice->SetNotifyWindowMessage( m_hWnd, WM_SPEAK, 0, 0 );
	//if ( !::CheckHr ( hr, "SetNotifyWindowMessage()", "Create" ) ) 
	//	return false;
	
	return true;
}

// ����Ҫ����������ɫ�����ԣ�ֻ֧�����ĺ�Ӣ��
bool CTTS::SetLanguage ( const DWORD dwLanguage )
{
	HRESULT hr;
	switch ( dwLanguage )
	{
		case SP_CHINESE:
			hr = SpFindBestToken( SPCAT_VOICES, L"language=804",
									NULL, &m_pToken );
			if ( !::CheckHr ( hr, "SpFindBestToken()", "SetLanguage()" ) ) 
				return false;
			
			hr = m_pVoice->SetVoice( m_pToken );
			if ( !::CheckHr ( hr, "SetVoice()", "SetLanguage()" ) ) 
				return false;
			break;

		case SP_ENGLISH:
			hr = SpFindBestToken( SPCAT_VOICES, L"language=409", NULL, &m_pToken );
			if ( !::CheckHr ( hr, "SpFindBestToken()����", "SetLanguage()" ) )  
				return false;
			
			hr = m_pVoice->SetVoice( m_pToken );
			if ( !::CheckHr ( hr, "SetVoice()����", "SetLanguage()" ) ) 
				return false;
			break;
		
		default:
			return false;
	}
	return true;
}

// ���ò��ŵ������� ��ΧΪ 0 �� 100
void CTTS::SetVolume ( const USHORT usVolume ) const
{
	if ( (usVolume > 100) || (usVolume<0)  ) {
		ShowMsg ( "SetVolume������ΧӦ�ô�0��100" );
		return;
	}
	m_pVoice->SetVolume( usVolume );
}

// �������
USHORT CTTS::GetVolume ()
{
	USHORT usVolume;
	m_pVoice->GetVolume ( &usVolume );
	return usVolume;
}

// ���������������ʣ� ��ΧΪ -10 �� 10
void CTTS::SetRate ( const LONG  RateAdjust ) const
{
	if ( (RateAdjust < -10) || (RateAdjust > 10) ) {
		ShowMsg ( "SetRate()���õ���ֵ������ -10��10֮��" );
		return;
	}
	m_pVoice->SetRate ( RateAdjust );
}

// ���ԭ�򲥷�����
LONG CTTS::GetRate ()
{
	LONG RateAdjust;
	m_pVoice->GetRate( &RateAdjust );
	return RateAdjust;
}

// ��һ��wav�ļ��󶨣���������ͨ�����첥�ţ�����������ļ�
bool CTTS::SetOutputToWav ( const WCHAR *pwchFileName, const SPSTREAMFORMAT FileFormat )
{
	HRESULT hr;
	CSpStreamFormat sOutputFormat;
	CComPtr<ISpStreamFormat>	cpOldStream;
	
	m_pVoice->GetOutputStream( &cpOldStream );
	//����wav�ļ���ʽ
	sOutputFormat.AssignFormat( FileFormat );
	
	hr = SPBindToFile( 	pwchFileName,
						SPFM_CREATE_ALWAYS,
						&m_pOutputStream,
						&sOutputFormat.FormatId(),
						sOutputFormat.WaveFormatExPtr() );
	if ( !::CheckHr ( hr, "SPBindToFile", "SetOutputToWav()" ) )  
		return false;
	
	hr = m_pVoice->SetOutput ( m_pOutputStream, TRUE );
	if ( !::CheckHr ( hr, "SetOutput()", "SetOutputToWav()" ) )  
		return false;
	
	return true;
}

// ȡ�����ļ��İ�
bool CTTS::UnSetOutputWithWav ()
{
	HRESULT hr;
	hr = m_pVoice->SetOutput ( m_pAudio, FALSE );
	if ( !::CheckHr ( hr, "SetOutput()", "UnSetOutputWithWav()" ) ) {
		m_pOutputStream->Close ();
		return false;
	}
	return true;
}

// ����
void CTTS::Speak ( const WCHAR *pwcs, const DWORD dwFlags ) const
{
	m_pVoice->Speak ( pwcs, dwFlags, NULL);
}

// ֹͣ����
void CTTS::Stop () const
{
	m_pVoice->Speak ( NULL, SPF_ASYNC, NULL );
}

// ��ͣ
void CTTS::Pause () const 
{
	m_pVoice->Pause ();
}

// ����ͣ����������
void CTTS::Resume () const
{
	m_pVoice->Resume ();
}







