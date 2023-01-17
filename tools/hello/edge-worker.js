'use strict';

const PINK_ENV = 'main--milo-college-pink--adobecom.hlx.live';
const OG_ENV = 'main--milo-college--adobecom.hlx.live';

const getFranklinReq = (url, request) => {
  const req = new Request(url, request);
  req.headers.set('x-forwarded-host', req.headers.get('host'));
  req.headers.set('x-byo-cdn-type', 'cloudflare');
  return req;
};

const handleRequest = async (request, env, ctx) => {
  const color = request.headers.get('x-adobe-floodgate');
  const resolvedEnv = color === 'pink' ? PINK_ENV : OG_ENV;

  const url = new URL(request.url);
  url.hostname = resolvedEnv;
  const req = getFranklinReq(url, request);
  // TODO: set the following header if push invalidation is configured
  // (see https://www.hlx.live/docs/setup-byo-cdn-push-invalidation#cloudflare)
  // req.headers.set('x-push-invalidation', 'enabled');
  let resp = await fetch(req);
  if (!resp.ok) {
    url.hostname = OG_ENV;
    const ogReq = getFranklinReq(url, request);
    resp = await fetch(ogReq);
  }

  resp = new Response(resp.body, resp);
  resp.headers.delete('age');
  resp.headers.delete('x-robots-tag');
  return resp;
};

export default {
  fetch: handleRequest,
};
