'use strict';

const PINK_ENV = 'main--milo-college-pink--adobecom.hlx.live';
const OG_ENV = 'main--milo-college--adobecom.hlx.live';

const getFranklinReq = (url, request) => {
  const req = new Request(url, request);
  req.headers.set('x-forwarded-host', req.headers.get('host'));
  req.headers.set('x-byo-cdn-type', 'cloudflare');
  return req;
};

const getFranklinResp = async (url, request, env) => {
  url.hostname = env;
  const req = getFranklinReq(url, request);
  return await fetch(req);
};

const handleRequest = async (request, env, ctx) => {
  const color = request.headers.get('x-adobe-floodgate');
  const url = new URL(request.url);
  let resp;

  if (color === 'pink') {
    resp = await getFranklinResp(url, request, PINK_ENV);
    if (!resp.ok) {
      // FG page does not exist. Fallback to page in regular tree.
      resp = await getFranklinResp(url, request, OG_ENV);
    }
  } else {
    resp = await getFranklinResp(url, request, OG_ENV);
  }

  resp = new Response(resp.body, resp);
  resp.headers.delete('age');
  resp.headers.delete('x-robots-tag');
  return resp;
};

export default {
  fetch: handleRequest,
};
