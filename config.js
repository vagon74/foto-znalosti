// Nastavení aplikace. Client ID doplňte po registraci aplikace v Microsoft Entra
// (viz README.md). Client ID není tajné – smí být v repozitáři.
window.FOTO_CONFIG = {
  clientId: '827ab626-9137-4573-b6fb-9bfec92abf3d',                     // Application (client) ID z Entra
  authority: 'https://login.microsoftonline.com/consumers', // osobní účet Microsoft
  scopes: ['Files.ReadWrite', 'offline_access'],
  inboxFolder: 'Znalosti/Inbox',      // cesta od kořene OneDrive
};
