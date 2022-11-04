
#include <windows.h>
#include <exdisp.h>
#include <exdispid.h>
#include <commctrl.h>
#pragma comment(lib,"Comctl32")
// Définir les messages pour notre application:
#define BEFORENAVIGATE2       WM_USER
#define DOWNLOADBEGIN         WM_USER+1
#define DOWNLOADCOMPLETE      WM_USER+2
#define NAVIGATECOMPLETE2     WM_USER+3
#define DOCUMENTCOMPLETE      WM_USER+4
#define COMMANDSTATECHANGE    WM_USER+5

//************************* Classe de gestion des événements *****************************
class Evenem : public IDispatch 
{
	private:
	long ref;
	HWND fenetre;
	BSTR url;

	public:
	Evenem(HWND fenet)
	{
	fenetre=fenet;
	}

	~Evenem()
	{
	SysFreeString(url);
	}

	STDMETHODIMP QueryInterface(REFIID iid, void ** ppvObject)
	{
	if (iid==IID_IUnknown || iid==IID_IDispatch || iid==DIID_DWebBrowserEvents2)
		{
			*ppvObject=this; 
			AddRef(); 
			return S_OK;
		} 
	else return E_NOINTERFACE;
	}

	ULONG STDMETHODCALLTYPE AddRef()
	{ 
	return InterlockedIncrement(&ref);  
	}

	ULONG STDMETHODCALLTYPE Release()
	{
	int tmp = InterlockedDecrement(&ref);
	if (tmp==0) delete this; 
	return tmp;
	}

	HRESULT STDMETHODCALLTYPE GetTypeInfoCount(unsigned int FAR* pctinfo)
	{ 
	return E_NOTIMPL;
	}

	HRESULT STDMETHODCALLTYPE GetTypeInfo(unsigned int iTInfo, LCID  lcid, ITypeInfo FAR* FAR*  ppTInfo)
	{
	return E_NOTIMPL;
	}

	HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, OLECHAR FAR* FAR* rgszNames, unsigned int cNames, LCID lcid, DISPID FAR* rgDispId)
	{ 
	return E_NOTIMPL; 
	}

	HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS FAR* pDispParams, VARIANT FAR* parResult, EXCEPINFO FAR* pExcepInfo, unsigned int FAR* puArgErr)
	{
	IUnknown *pIUnk;
	VARIANT  *vurl ;

	if (!pDispParams) return E_INVALIDARG;
  
	switch (dispIdMember)
		{
		case DISPID_BEFORENAVIGATE2:
			// Détérminer l'objet courant:
			pIUnk=pDispParams->rgvarg[6].pdispVal;
			// Envoyer le message BEFORENAVIGATE2 à notre fenêtre:
			SendMessage(fenetre,BEFORENAVIGATE2,(WPARAM)pIUnk,0);
			break;

		case DISPID_DOWNLOADBEGIN:
			// Envoyer le message DOWNLOADBEGIN à notre fenêtre:
			SendMessage(fenetre,DOWNLOADBEGIN,0,0);
			break;

		case DISPID_DOWNLOADCOMPLETE:
			// Envoyer le message DOWNLOADCOMPLETE à notre fenêtre:
			SendMessage(fenetre,DOWNLOADCOMPLETE,0,0);
			break;

		case DISPID_NAVIGATECOMPLETE2:
			// Déterminer l'objet courant:
			pIUnk=pDispParams->rgvarg[1].pdispVal;
			// Récupérer l'URL courante:
			vurl= pDispParams->rgvarg[0].pvarVal;
			url = vurl->bstrVal;
			// Envoyer le message NAVIGATECOMPLETE2 à notre fenêtre:
			SendMessage(fenetre,NAVIGATECOMPLETE2,(WPARAM)pIUnk,(LPARAM)url);
			break;

		case DISPID_DOCUMENTCOMPLETE:
			// Déterminer l'objet courant:
			pIUnk=pDispParams->rgvarg[1].pdispVal;
			// Envoyer le message DOCUMENTCOMPLETE à notre fenêtre:
			SendMessage(fenetre,DOCUMENTCOMPLETE,(WPARAM)pIUnk,0);
			break;
			
		case DISPID_COMMANDSTATECHANGE:
			// Déterminer la commande dont l'état a changé:
			long command;
			command =pDispParams->rgvarg[1].lVal;
			// Déterminer le nouvel état de cette commande:
			VARIANT_BOOL etat;
			etat=pDispParams->rgvarg[0].boolVal;
			// Envoyer le message COMMANDESTATECHANGE à notre fenêtre:
			SendMessage(fenetre,COMMANDSTATECHANGE,(WPARAM)command,(LPARAM)etat);
			break;

		case DISPID_NEWWINDOW2:
			// Bloquer toutes les fenêtres popup:
			pDispParams->rgvarg[0].pvarVal->vt = VT_BOOL;
			pDispParams->rgvarg[0].pvarVal->boolVal = VARIANT_TRUE;
			break;

		default:
			break;
		}	
  return S_OK;
	}
};
//***********************************************************************************

// Variables globales:
IWebBrowser2   *pIWeb;
WNDPROC OldEditProc;
HWND hConteneur,hAdresse;


/******** Procédure de sous-classement de l'EDIT de la barre d'adresse ********/
LRESULT CALLBACK EditProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message) 
	{
	case WM_CHAR: 
			// Empêcher les beep après appui sur ENTREE:
			if (wParam == VK_RETURN) return 0;
			break;

	case WM_KEYDOWN:
		// Si la touche tapée est "ENTREE":
		if (wParam == VK_RETURN)
		{
			// Récupérer et convertir l'URL de la barre d'adresse:
			WCHAR url[256];
			char buff[256];
			GetWindowText(hAdresse,buff,256);
			MultiByteToWideChar (CP_ACP, 0,buff, -1, url, 256);
			// Lancer la navigation:
			pIWeb->Navigate(url,0,0,0,0);
			return 0;
		}
		break;

	default:
		break;
	}
	// Appeler la procédure originale:
	return CallWindowProc(OldEditProc, hwnd, message, wParam, lParam);
}

LRESULT CALLBACK WndProc( HWND hWnd, UINT messg, WPARAM wParam, LPARAM lParam )
{
	static HWND hPrecedente,hSuivante,hArreter,hActualiser,hDemarrage,hEnregistrer,hImprimer,hAller;
	static HWND hCadre1,hCadre2,hTadresse,hEtat;
	static int PageCounter=0;
	static int ObjCounter=0;
	char tampon[256];
	static BSTR titre;
	WCHAR url[256];

	switch(messg)
	{
		case WM_CREATE:
			// Création de tous les contrôles:
			hPrecedente=CreateWindow("BUTTON","Précédente",WS_CHILD | WS_VISIBLE,5,6,90,20,hWnd,0,0,0);
			hSuivante=CreateWindow("BUTTON","Suivante",WS_CHILD | WS_VISIBLE,105,6,90,20,hWnd,0,0,0);
			hArreter=CreateWindow("BUTTON","Arrêter",WS_CHILD | WS_VISIBLE,205,6,90,20,hWnd,0,0,0);
			hActualiser=CreateWindow("BUTTON","Actualiser",WS_CHILD | WS_VISIBLE,305,6,90,20,hWnd,0,0,0);
			hDemarrage=CreateWindow("BUTTON","Démarrage",WS_CHILD | WS_VISIBLE,405,6,90,20,hWnd,0,0,0);
			hEnregistrer=CreateWindow("BUTTON","Enregistrer",WS_CHILD | WS_VISIBLE,505,6,90,20,hWnd,0,0,0);
			hImprimer=CreateWindow("BUTTON","Imprimer",WS_CHILD | WS_VISIBLE,605,6,90,20,hWnd,0,0,0);
			hTadresse=CreateWindow("STATIC","Adresse  :",WS_CHILD | WS_VISIBLE | SS_CENTER,7,35,90,20,hWnd,0,0,0);
			hAller=CreateWindow("BUTTON","Aller",WS_CHILD | WS_VISIBLE,610,34,80,20,hWnd,0,0,0);
			hCadre1=CreateWindow("BUTTON",0,WS_CHILD | WS_VISIBLE | BS_GROUPBOX,05,22,690,36,hWnd,0,0,0);
			hAdresse=CreateWindowEx(WS_EX_CLIENTEDGE,"EDIT",0,WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL ,105,34,490,20,hWnd,0,0,0);
			// Obtenir l'adresse de la procédure de l'EDIT de la barre d'adresse pour le sous-classer:
			hEtat=CreateStatusWindow(WS_CHILD|WS_VISIBLE,"ProjetTosca",hWnd,6000);
			OldEditProc= (WNDPROC) SetWindowLong(hAdresse, GWLP_WNDPROC, (LPARAM)EditProc);
			// Griser tous les boutons:
			EnableWindow(hPrecedente,0);
			EnableWindow(hSuivante,0);
			EnableWindow(hArreter,0);
			EnableWindow(hActualiser,0);
			EnableWindow(hEnregistrer,0);
			EnableWindow(hImprimer,0);
			EnableWindow(hAller,0);
			break;
		
		case WM_COMMAND:
			// Griser le bouton "Aller" si la barre d'adresse est vide:
			GetWindowText(hAdresse,tampon,256);
			EnableWindow(hAller,lstrlen(tampon));

			if( (HWND)lParam == hPrecedente ) // Clic sur "Précédente"
			{
				// Aller à la page précédente:
				pIWeb->GoBack();
				break;
			}
			if( (HWND)lParam == hSuivante ) // Clic sur "Suivante"
			{
				// Aller à la page suivante:
				pIWeb->GoForward();
				break;
			}
			if( (HWND)lParam == hArreter ) // Clic sur "Arrêter"
			{
				// Initialiser les compteurs:
				PageCounter=ObjCounter=0;
				// Arrêter la navigation:
				pIWeb->Stop();
				// Griser le bouton "Arrêter" et dégriser "Actualiser"
				EnableWindow(hArreter,0);
				EnableWindow(hActualiser,1);
				SetWindowText(hEtat,"Arrêté");
				break;
			}
			if((HWND) lParam == hActualiser ) // Clic sur "Actualiser"
			{
				// Actualier la page:
				pIWeb->Refresh2(0);
				break;
			}
			if( (HWND)lParam == hDemarrage ) // Clic sur "Démarrage"
			{
				// Lancer la page de démarrage:
							MultiByteToWideChar (CP_ACP, 0,"https://services.csvdc.qc.ca/toscanetcrif/asp/Tosca.aspx", -1, url, 256);
			// Lancer la navigation:
			pIWeb->Navigate(url,0,0,0,0);

			//	pIWeb->GoHome();
				break;
			}
			if((HWND) lParam == hEnregistrer ) // Clic sur "Enregistrer"
			{
				// Enregister la page:
				pIWeb->ExecWB(OLECMDID_SAVEAS,OLECMDEXECOPT_DODEFAULT, 0,0);
				break;
			}
			if( (HWND)lParam == hImprimer ) //Clic sur "Imprimer"
			{
				// Imprimer la page:
				pIWeb->ExecWB(OLECMDID_PRINT,OLECMDEXECOPT_DONTPROMPTUSER, 0,0);
				break;
			}

			if((HWND) lParam == hAller ) // Clic sur "Aller"
			{
			// Récupérer et convertir le lien de la barre d'adresse:
			WCHAR url[256];
			GetWindowText(hAdresse,tampon,256);
			MultiByteToWideChar (CP_ACP, 0,tampon, -1, url, 256);
			// Lancer la navigation:
			pIWeb->Navigate(url,0,0,0,0);
			}
			break;

		case BEFORENAVIGATE2:
			// Si c'est le début d'une nouvelle page alors initialier les compteurs:
			if ((IUnknown*)wParam==pIWeb)	PageCounter=ObjCounter=0;
			// Incrémenter le compteur de pages:
			PageCounter++;
			// Dégriser le bouton "Arrêter" et griser les autres:
			EnableWindow(hArreter,1);
			EnableWindow(hActualiser,0);
			EnableWindow(hEnregistrer,0);
			EnableWindow(hImprimer,0);
			break;

		case DOWNLOADBEGIN :
			// Incrémenter le compteur d'objets:
			ObjCounter++;
			// Dégriser le bouton "Arrêter" et griser les autres:
			EnableWindow(hArreter,1);
			EnableWindow(hActualiser,0);
			EnableWindow(hEnregistrer,0);
			EnableWindow(hImprimer,0);
			// Afficher "Navigation" dans le cadre d'état:
			SetWindowText(hEtat,"Navigation");
			break;

		case DOWNLOADCOMPLETE :
			// Décrémenter le compteur d'objets:
			ObjCounter--;
			// Si les deux compteurs sonts nuls:
			if (PageCounter==0 && ObjCounter==0)
			{
				// Griser le bouton "Arrêter" et dégriser les autres:
				EnableWindow(hArreter,0);
				EnableWindow(hActualiser,1);
				EnableWindow(hEnregistrer,1);
				EnableWindow(hImprimer,1);
				// Afficher "Terminé" dans le cadre d'état:
				SetWindowText(hEtat,"Terminé");
			}
			break;

		case DOCUMENTCOMPLETE:
			// Décrémenter le compteur de pages:
			PageCounter--;
			// Si le compteur de pages est nul:
			if (PageCounter==0)
			{	
				// Obtenir, convertir et afficher le titre de la page sur la barre de titre:
				pIWeb->get_LocationName(&titre);
				WideCharToMultiByte(CP_ACP,0,titre,-1,tampon,256,0,0);
				lstrcat(tampon," - ProjetTosca");
				SetWindowText(hWnd,tampon);
				// Griser le bouton "Arrêtter" et dégriser les autres:
				EnableWindow(hArreter,0);
				EnableWindow(hActualiser,1);
				EnableWindow(hEnregistrer,1);
				EnableWindow(hImprimer,1);
				// Afficher "Terminé" dans le cadre d'état:
				SetWindowText(hEtat,"Terminé");
			}
			break;

		case NAVIGATECOMPLETE2:
			// Si le texte html de la page est chargé:
			if((IUnknown*)wParam==pIWeb)
			{
				// Convertir et afficher le lien sur la barre d'adresses:
				WideCharToMultiByte(CP_ACP,0,(LPCWSTR)lParam,-1,tampon,256,0,0);
				SetWindowText(hAdresse,tampon);
				// Obtenir, convertir et afficher le titre de la page sur la barre de titre:
				pIWeb->get_LocationName(&titre);
				WideCharToMultiByte(CP_ACP,0,titre,-1,tampon,256,0,0);
				lstrcat(tampon," - ProjetTosca");
				SetWindowText(hWnd,tampon);
			}
			break;

		case COMMANDSTATECHANGE:
			// Détecter et changer l'état des boutons "Précédente" et "Suivante":
			if (wParam==2) EnableWindow(hPrecedente,(BOOL)lParam);
			if (wParam==1) EnableWindow(hSuivante,(BOOL)lParam);
			break;

		case WM_SIZE:{
			RECT rc;
			MoveWindow(hConteneur,5,62,LOWORD(lParam)-8, HIWORD(lParam)-100,1);
			GetClientRect(hConteneur,&rc);
			// Redimensionner le conteneur quand la taille de la fenêtre change:
			MoveWindow(hEtat,0,rc.bottom-32,rc.right, 25,1);
					 }return 0;

		case WM_CLOSE:
			// Libérer le BSTR 
			SysFreeString(titre);
			// Détruire la fenêtre principale:
			DestroyWindow(hWnd);
			break;

		case WM_DESTROY:
			// Envoyer le message de sortie du programme:
			PostQuitMessage( 0 );
			break;
	
		default:
			// Retour:
			return( DefWindowProc( hWnd, messg, wParam, lParam ) );
	}
	return 0;
}
/***************************************************************************/

/********************* Fonction WinMan ************************************************/
int WINAPI WinMain( HINSTANCE hInst,HINSTANCE hPreInst,LPSTR lpszCmdLine, int nCmdShow )
{
	// Déclarer notre classe de fenêtre et définir ses membres:
	WNDCLASS wc;
	char NomClasse[]   = "ProjetTosca";
	wc.lpszClassName 	= NomClasse;
	wc.hInstance 		= hInst;
	wc.lpfnWndProc		= WndProc;
	wc.hCursor			= LoadCursor( NULL, IDC_ARROW );
	wc.hIcon			= LoadIcon( wc.hInstance,(LPCSTR) 101 );
	wc.lpszMenuName	    = NULL;
	wc.hbrBackground	= GetSysColorBrush(COLOR_BTNFACE);;
	wc.style			= 0;
	wc.cbClsExtra		= 0;
	wc.cbWndExtra		= 0;
	if (!RegisterClass(&wc)) return 0;
	RECT rc;
	GetClientRect(GetDesktopWindow(), &rc);
	HWND hWnd = CreateWindow( NomClasse,"ProjetTosca", WS_POPUP |WS_BORDER | WS_SYSMENU|WS_CAPTION,rc.left,rc.top,rc.right,rc.bottom, 0, 0, hInst,0);
	ShowWindow(hWnd, nCmdShow );
	UpdateWindow( hWnd );
	HINSTANCE hDLL = LoadLibrary("atl.dll");
	typedef HRESULT (WINAPI *PAttachControl)(IUnknown*, HWND,IUnknown**);
	PAttachControl AtlAxAttachControl = (PAttachControl) GetProcAddress(hDLL, "AtlAxAttachControl");
	RECT rect;
	GetClientRect(hWnd,&rect);
	hConteneur=CreateWindowEx(WS_EX_CLIENTEDGE,"BUTTON","",WS_CHILD | WS_VISIBLE,5,62,rect.right-8,rect.bottom-67,hWnd,0,0,0);
	CoInitialize(0);
CoCreateInstance(CLSID_WebBrowser,0,CLSCTX_ALL,IID_IWebBrowser2,(void**)&pIWeb);
	AtlAxAttachControl(pIWeb,hConteneur,NULL);
	IConnectionPointContainer* pCPContainer;
    pIWeb->QueryInterface(IID_IConnectionPointContainer,(void**)&pCPContainer);
	IConnectionPoint *pConnectionPoint;
	pCPContainer->FindConnectionPoint(DIID_DWebBrowserEvents2, &pConnectionPoint);
	Evenem *pEvnm;
	pEvnm= new Evenem(hWnd);
	DWORD dwCookie = 0;
	pConnectionPoint->Advise(pEvnm, &dwCookie);
	if (pCPContainer) pCPContainer->Release();
		WCHAR url[256];
						MultiByteToWideChar (CP_ACP, 0,"https://services.csvdc.qc.ca/toscanetcrif/asp/Tosca.aspx", -1, url, 256);
			pIWeb->Navigate(url,0,0,0,0);
	MSG Msg;
	while( GetMessage( &Msg, 0, 0, 0 ) )
	{
		TranslateMessage( &Msg );
		DispatchMessage( &Msg );
	}
	pConnectionPoint->Unadvise(dwCookie);
   pConnectionPoint->Release();
	delete pEvnm;
	pIWeb->Release();
    CoUninitialize();		
	FreeLibrary(hDLL);
	return( Msg.wParam);
}
