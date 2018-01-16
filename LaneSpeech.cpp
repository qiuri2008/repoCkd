#include "stdafx.h"

#define USE_SPEECH_DLL
#define LANE_SPEECH_EXPORTS

#include "LaneSpeech.h"

#ifdef USE_SPEECH_DLL	//定义了USE_SPEECH_DLL，就按生成DLL，声明导出导入类
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

// 建立一个SR引擎， dwLanguage为语言种类，bInproc为 本识别引擎是否为本程序独享
bool CSR::Create ( const DWORD dwLanguage,
				   const bool bInproc )
{
	HRESULT hr;

	if ( bInproc ) {//独享的引擎
		hr = m_pSREngine.CoCreateInstance ( CLSID_SpInprocRecognizer );
		if ( !::CheckHr(hr, "CoCreateInstance()", "Create()") ) 
			return false; 
		
		hr = SpCreateDefaultObjectFromCategoryId ( SPCAT_AUDIOIN, &m_pAudio );	//建立一个默认音频流
		if ( !::CheckHr(hr, "SpCreateDefaultObjectFromCategoryId()", "Create()") ) 
			return false; 
		
		hr = m_pSREngine->SetInput ( m_pAudio, TRUE );	//为识别引擎设置音频输入
		if ( !::CheckHr(hr, "SetInput()", "Create()") ) 
			return false; 
	}
	else {//共享的引擎
		hr = m_pSREngine.CoCreateInstance ( CLSID_SpSharedRecognizer );//建立一个共享的SR引擎。
		if ( !::CheckHr(hr, "CoCreateInstance", "Create()") ) 
			return false; 
	}
	
	if ( dwLanguage == SP_CHINESE )  //识别语言为中文
		hr = SpFindBestToken ( SPCAT_RECOGNIZERS, L"language=804", 
								NULL, &m_pToken );
	else	//识别语言为中文
		hr = SpFindBestToken ( SPCAT_RECOGNIZERS, L"language=409", 
								NULL, &m_pToken );
	if ( !::CheckHr(hr, "SpFindBestToken()", "Create()") )
		return false;
	
	hr = m_pSREngine->SetRecognizer ( m_pToken );	//设置设别语言
	if ( !::CheckHr(hr, "SetRecognizer()", "Create()") )
		return false;

	hr = m_pSREngine->CreateRecoContext ( &m_pSRContext );	//建立上下文
	if ( !::CheckHr(hr, "CreateRecoContext()", "Create()") )
		return false;

	hr = m_pSRContext->SetNotifyWin32Event();			//设置win32事件
	if ( !::CheckHr(hr, "SetNotifyWin32Event()", "Create()") )
		return false;

	//设置要处理的事件
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

// 加载识别引擎要用的的文法文件
bool CSR::LoadCrammarFile ( const WCHAR *pwchFilename )
{
	HRESULT hr;
	
	hr = m_pSRContext->CreateGrammar ( ullGrammerID, &m_pSRGrammar );	//建立文法接口
	if ( !::CheckHr(hr, "CreateGrammar()", "LoadCrammarFile()") )
		return false;
	
	hr = m_pSRGrammar->LoadCmdFromFile ( pwchFilename, SPLO_DYNAMIC );	//加载文法文件
	if ( !::CheckHr(hr, "LoadCmdFromFile()", "LoadCrammarFile()") )
		return false;

	return true;
}

// 把识别引擎设置为从wav文件获得输入，pwchFileName为文件名，FileFormat为文件格式
bool CSR::SetInputFromWav ( const WCHAR *pwchFileName, const SPSTREAMFORMAT FileFormat )
{
	HRESULT hr;

	hr = m_pSRGrammar->SetRuleState ( NULL, NULL, SPRS_INACTIVE );	//取消规则激活
	if ( !::CheckHr( hr, "SetRuleState()", "SetInputFromWav()" ) )  
		return false;
	
	if ( m_pInputStream ) {
		m_pInputStream.Release();
		m_pInputStream = NULL;
	}
	
	hr = m_pInputStream.CoCreateInstance(CLSID_SpStream);			//建立流接口
	if ( !::CheckHr( hr, "CoCreateInstance()", "SetInputFromWav()" ) )  
		return false;
	
	CSpStreamFormat sInputFormat;
	//hr = sInputFormat.AssignFormat(SPSF_22kHz16BitStereo);			//设置wav文件格式
	hr = sInputFormat.AssignFormat ( FileFormat );
	if ( !::CheckHr( hr, "AssignFormat()", "SetInputFromWav()" ) ) 
		return false;
	
	hr = m_pInputStream->BindToFile	(	pwchFileName,				//绑定wav文件到流
										SPFM_OPEN_READONLY,
										&sInputFormat.FormatId(),
										sInputFormat.WaveFormatExPtr(),
										SPFEI_ALL_EVENTS  );
	if ( !::CheckHr( hr, "BindToFile()", "SetInputFromWav()" ) ) 	
		return false;
	
	hr = m_pSREngine->SetInput ( m_pInputStream, true );			//设置引擎的输入
	if ( !::CheckHr( hr, "SetInput()", "SetInputFromWav()" ) ) 
		return false;
	
	//设置要处理的事件
	hr = m_pSRContext->SetInterest (SPFEI(SPEI_RECOGNITION) | SPFEI( SPEI_FALSE_RECOGNITION ) |
									SPFEI(SPEI_END_SR_STREAM) | SPFEI(SPEI_START_SR_STREAM) |
									SPFEI(SPEI_SOUND_END) | SPFEI( SPEI_SOUND_START ),
									SPFEI(SPEI_RECOGNITION) | SPFEI( SPEI_FALSE_RECOGNITION ) |
									SPFEI(SPEI_END_SR_STREAM) | SPFEI(SPEI_START_SR_STREAM) |
									SPFEI(SPEI_SOUND_END) | SPFEI( SPEI_SOUND_START ));
	
	return true;
}

// 识别主的处理过程，识别引擎识别短语的过程就放在这里，参数为超时值
void CSR::ProcessReco( const DWORD dwMilliseconds )
{
	//当有事件发生，继续处理，timeout应该设置的比audio流长，
	//或者期望传递一个合理的时间总数在不再识别以前，
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
					//ShowMsg("识别成功");
					m_bRecoSuccess = true;
					OnRecoSuccess ( spEvent.RecoResult() );
					//m_bRecoEnd	   = true;
					break;
				case SPEI_FALSE_RECOGNITION:
					//ShowMsg("识别失败");
					m_bRecoSuccess = false;
					//m_bRecoEnd	   = true;
					break;
				
				case SPEI_END_SR_STREAM:
					m_bRecoEnd			= true;
					fEndStreamReached	= true;
					//ShowMsg("识别结束");
					bTimeOut = false;
					break;

				case SPEI_START_SR_STREAM:
					//ShowMsg("识别开始");
					m_bRecoEnd	   = false;
					bTimeOut = true;
					break;

				//case SPEI_SOUND_START:
				//	m_bRecoEnd			= true;
					//fEndStreamReached	= true;
				//	ShowMsg("声音开始");
				//	break;

				//case SPEI_SOUND_END:
				//	m_bRecoEnd			= true;
					//fEndStreamReached	= true;
				//	ShowMsg("声音结束");
				//	break;
			}
			spEvent.Clear();
		}
	}
	if ( bTimeOut || !fEndStreamReached ) {
		m_bRecoEnd	= true;
	}

}

// 识别成功的情况下，就保存识别的短语
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
		if ( pPhrase ) {  // 释放短语对象。
			::CoTaskMemFree(pPhrase);
		}
	}
}

// 返回识别是否结束
bool CSR::IsRecoEnd() const
{
	return	m_bRecoEnd;
}

// 返回识别是否成功
bool CSR::IsRecoSuccess() const
{
	return	m_bRecoSuccess;
}

// 开始识别
bool CSR::StartReco()
{
	HRESULT hr;
	hr = m_pSRGrammar->SetRuleState ( NULL, NULL, SPRS_ACTIVE );	//激活所有规则
	if ( !::CheckHr( hr, "SetRuleState()", "SetInputFromWav()" ) )  
		return false;
	return true;
}

// 获得识别的短语， UNICODE 字符版本
void CSR::GetRecoWordW ( wchar_t * pwchRecoWord )
{
	wcscpy ( pwchRecoWord, m_wchRecoWord );
}

// 获得识别的短语， ANSI 字符版本
void CSR::GetRecoWordA (char * pchRecoWord )
{
	char temp[128];
	WideCharToMultiByte ( CP_ACP, //把char转换成wchar_t
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
//	CTTS 类
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

// 建立一个CTTS引擎
bool CTTS::Create( const DWORD dwLanguage )
{
	HRESULT hr;
	if ( m_pVoice )
		return true;
	hr = m_pVoice.CoCreateInstance ( CLSID_SpVoice );
	if ( !::CheckHr( hr, "CoCreateInstance()", "Create" ) )  
		return false;
	
	SetLanguage ( dwLanguage );
	
	hr = SpCreateDefaultObjectFromCategoryId ( SPCAT_AUDIOIN, &m_pAudio );//建立一个默认音频流
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

// 设置要播放语音角色的语言，只支持中文和英文
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
			if ( !::CheckHr ( hr, "SpFindBestToken()错误", "SetLanguage()" ) )  
				return false;
			
			hr = m_pVoice->SetVoice( m_pToken );
			if ( !::CheckHr ( hr, "SetVoice()错误", "SetLanguage()" ) ) 
				return false;
			break;
		
		default:
			return false;
	}
	return true;
}

// 设置播放的音量， 范围为 0 到 100
void CTTS::SetVolume ( const USHORT usVolume ) const
{
	if ( (usVolume > 100) || (usVolume<0)  ) {
		ShowMsg ( "SetVolume音量范围应该从0到100" );
		return;
	}
	m_pVoice->SetVolume( usVolume );
}

// 获得音量
USHORT CTTS::GetVolume ()
{
	USHORT usVolume;
	m_pVoice->GetVolume ( &usVolume );
	return usVolume;
}

// 设置语音播放速率， 范围为 -10 到 10
void CTTS::SetRate ( const LONG  RateAdjust ) const
{
	if ( (RateAdjust < -10) || (RateAdjust > 10) ) {
		ShowMsg ( "SetRate()设置的数值必须在 -10到10之间" );
		return;
	}
	m_pVoice->SetRate ( RateAdjust );
}

// 获得原因播放速率
LONG CTTS::GetRate ()
{
	LONG RateAdjust;
	m_pVoice->GetRate( &RateAdjust );
	return RateAdjust;
}

// 与一个wav文件绑定，语音将不通过外响播放，而是输出到文件
bool CTTS::SetOutputToWav ( const WCHAR *pwchFileName, const SPSTREAMFORMAT FileFormat )
{
	HRESULT hr;
	CSpStreamFormat sOutputFormat;
	CComPtr<ISpStreamFormat>	cpOldStream;
	
	m_pVoice->GetOutputStream( &cpOldStream );
	//设置wav文件格式
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

// 取消与文件的绑定
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

// 播放
void CTTS::Speak ( const WCHAR *pwcs, const DWORD dwFlags ) const
{
	m_pVoice->Speak ( pwcs, dwFlags, NULL);
}

// 停止播放
void CTTS::Stop () const
{
	m_pVoice->Speak ( NULL, SPF_ASYNC, NULL );
}

// 暂停
void CTTS::Pause () const 
{
	m_pVoice->Pause ();
}

// 从暂停处继续播放
void CTTS::Resume () const
{
	m_pVoice->Resume ();
}







