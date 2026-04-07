function normalizePrivateKey(value) {
  let key = String(value || '').trim();

  if (
    (key.startsWith('"') && key.endsWith('"')) ||
    (key.startsWith("'") && key.endsWith("'"))
  ) {
    key = key.slice(1, -1);
  }

  return key.replace(/\\n/g, '\n');
}

function getEnvDebugInfo() {
  const serviceAccountJson = process.env.GOOGLE_SERVICE_ACCOUNT_JSON || '';
  const clientEmail = process.env.GOOGLE_CLIENT_EMAIL || '';
  const privateKey = process.env.GOOGLE_PRIVATE_KEY || '';
  const normalizedPrivateKey = normalizePrivateKey(privateKey);

  return {
    vercel: Boolean(process.env.VERCEL),
    hasServiceAccountJson: Boolean(serviceAccountJson.trim()),
    serviceAccountJsonLength: serviceAccountJson.length,
    serviceAccountJsonLooksJson: serviceAccountJson.trim().startsWith('{'),
    hasClientEmail: Boolean(clientEmail.trim()),
    clientEmailLength: clientEmail.length,
    clientEmailLooksLikeServiceAccount: clientEmail.includes('iam.gserviceaccount.com'),
    hasPrivateKey: Boolean(privateKey.trim()),
    privateKeyLength: privateKey.length,
    privateKeyHasBegin: normalizedPrivateKey.includes('-----BEGIN PRIVATE KEY-----'),
    privateKeyHasEnd: normalizedPrivateKey.includes('-----END PRIVATE KEY-----'),
    privateKeyHasRealNewlines: normalizedPrivateKey.includes('\n'),
    usingCredentialMode: serviceAccountJson.trim()
      ? 'GOOGLE_SERVICE_ACCOUNT_JSON'
      : clientEmail.trim() && privateKey.trim()
        ? 'GOOGLE_CLIENT_EMAIL/GOOGLE_PRIVATE_KEY'
        : 'none'
  };
}

module.exports = (req, res) => {
  res.status(200).json(getEnvDebugInfo());
};
