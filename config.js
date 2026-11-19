// Public settings for this deployment. Nothing here is secret.
// clientId / tenantId come from IT's app registration in Microsoft Entra ID.
// They can also be entered in the app (Settings), which overrides these values.
export const CONFIG = {
  clientId: '',
  tenantId: '',
  // Delegated Microsoft Graph permissions. The app can only reach the signed-in user's own mailbox.
  scopes: ['openid', 'profile', 'offline_access', 'User.Read', 'Mail.ReadWrite', 'Contacts.Read', 'People.Read'],
};
