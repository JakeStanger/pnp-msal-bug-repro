import { sp } from '@pnp/sp-commonjs/presets/all';
import { MsalFetchClient } from '@pnp/nodejs-commonjs';
import fs from 'fs';
import dotenv from 'dotenv';

dotenv.config();

async function getSpClient() {
  const spClient = await sp.createIsolated();
  const privateKey = getPrivateKey();

  const env = process.env;

  sp.setup({
    sp: {
      baseUrl: `https://${env.SP_TENANT_SUBDOMAIN}.sharepoint.com/${env.SP_SITE}`,
      fetchClientFactory: () => {
        return new MsalFetchClient(
          {
            auth: {
              authority: `https://login.microsoftonline.com/${env.AAD_TENANT_ID}/`,
              clientCertificate: {
                thumbprint: env.AAD_CERT_THUMB,
                privateKey,
              },
              clientId: env.AAD_CLIENT_ID,
            },
          },
          [`https://${env.SP_TENANT_SUBDOMAIN}.sharepoint.com/.default`]
        ); // you must set the scope for SharePoint access
      },
    },
  });

  return spClient;
}

function getPrivateKey() {
  return fs.readFileSync(process.env.AAD_CERT_KEY_PATH).toString('utf-8');
}

async function run() {
  const spClient = await getSpClient();

  const lists = await spClient.web.lists.select('Title').get();
  console.log(lists);

  const catalog = await spClient.web.getAppCatalog();
  const apps = await catalog.get();
  console.log(apps);
}

run().catch(console.error);