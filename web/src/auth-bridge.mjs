import {broadcastResponseToMainFrame} from '@azure/msal-browser/redirect-bridge';
broadcastResponseToMainFrame().catch(()=>{document.body.textContent='Microsoft bağlantısı ana pencereye aktarılamadı. Ana sayfada Hesap Ekle düğmesini yeniden kullanın.';});
