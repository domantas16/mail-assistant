// Public settings for this deployment. Nothing here is secret.
// clientId / tenantId come from IT's app registration in Microsoft Entra ID.
// They can also be entered in the app (Settings), which overrides these values.
export const CONFIG = {
  clientId: '',
  tenantId: '',
  // Delegated Microsoft Graph permissions: the signed-in user's own mailbox, plus shared mailboxes
  // that user already has access to in Exchange. The app can never reach more than the user can.
  scopes: ['openid', 'profile', 'offline_access', 'User.Read', 'Mail.ReadWrite', 'Mail.Send', 'Mail.ReadWrite.Shared', 'Mail.Send.Shared', 'Contacts.Read', 'People.Read'],
};
