import {PublicClientApplication,InteractionRequiredAuthError} from '@azure/msal-browser';
export const SCOPES=['User.Read','Mail.Read'];
export const validClientId=value=>/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value||'')&&!/^0{8}-0{4}-0{4}-0{4}-0{12}$/.test(value);
export const redirectUri=()=>new URL('auth.html',new URL(import.meta.env.BASE_URL,window.location.origin)).href;
export class MicrosoftAuth{
 constructor(clientId){if(!validClientId(clientId))throw new Error('Microsoft uygulama kimliği henüz yapılandırılmadı.');this.client=new PublicClientApplication({auth:{clientId,authority:'https://login.microsoftonline.com/common',redirectUri:redirectUri()},cache:{cacheLocation:'memoryStorage'},system:{loggerOptions:{piiLoggingEnabled:false,loggerCallback:()=>{}}}});this.ready=this.client.initialize();}
 async connect(){await this.ready;return this.client.loginPopup({scopes:SCOPES,prompt:'select_account',redirectUri:redirectUri()});}
 async token(account){await this.ready;try{return (await this.client.acquireTokenSilent({account:account.auth,scopes:SCOPES,redirectUri:redirectUri()})).accessToken;}catch(error){if(error instanceof InteractionRequiredAuthError)throw new Error('Microsoft oturumu yenilenmeli. Hesap Ekle ile bu hesabı yeniden bağlayın.');throw new Error('Microsoft erişim belirteci alınamadı. Hesabı yeniden bağlayın.');}}
 async disconnect(account){await this.ready;await this.client.clearCache({account:account.auth});}
}
export function authError(error){
 if(error.errorCode==='user_cancelled')return 'Microsoft giriş penceresi kapatıldı.';
 if(error.errorCode==='popup_window_error'||error.errorCode==='empty_window_error')return 'Microsoft giriş penceresi açılamadı. Bu sayfa için açılır pencerelere izin verin.';
 const details=String(error.errorMessage||error.message||'');
 if(details.includes('AADSTS50011'))return 'Microsoft uygulamasının web dönüş adresi eşleşmiyor. Bağlantı kurulumu bölümündeki adres Microsoft Entra’da SPA olarak tanımlanmalıdır.';
 if(details.includes('AADSTS65001')||details.includes('consent_required'))return 'Microsoft kurum yöneticinizin posta okuma iznini onaylaması gerekiyor.';
 return 'Microsoft bağlantısı tamamlanamadı. Uygulama kimliği ve web dönüş adresini kontrol edin.';
}
