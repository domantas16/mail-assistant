// Public settings for this deployment. Nothing here is secret.
// clientId / tenantId come from IT's app registration in Microsoft Entra ID.
// They can also be entered in the app (Settings), which overrides these values.
export const CONFIG = {
  clientId: '',
  tenantId: '',
  // Delegated Microsoft Graph permissions. The app starts read-only with the smallest set:
  // sign-in plus reading the signed-in user's own mail. It can never reach more than the user can.
  baseScopes: ['openid', 'profile', 'offline_access', 'User.Read', 'Mail.Read'],
  // Optional features, each switched on in Settings. Microsoft asks for the extra
  // permission (or IT approves it) only when a feature is switched on.
  extraScopes: {
    send: ['Mail.ReadWrite', 'Mail.Send'], // send and schedule replies from the app, mark as read
    sharedRead: ['Mail.Read.Shared'], // read shared mailboxes (info@, sales@ ...)
    sharedSend: ['Mail.ReadWrite.Shared', 'Mail.Send.Shared'], // reply from shared mailboxes
    contacts: ['Contacts.Read', 'People.Read'], // Outlook contacts and people search
  },
};
