# PnPJS MSAL/App Catalog Bug Reproduction

Repo to demonstrate the presumed bug at <https://github.com/pnp/pnpjs/issues/1719>.

## Steps

1. Run `yarn` to install dependencies.
2. Copy `.env.example` to `.env` and fill out the values for MSAL certificate authentication.
3. Run `yarn start:dev` to start using `ts-node` or run `yarn build && yarn start` to run using `node`.

The program will:
- Get an isolated SP client (important!)
- Fetch and log all lists for the site
- (Attempt to) Fetch and log all apps for the site