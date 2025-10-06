# One Page Report Generator

A zero-backend tool that turns Smartsheet trackers and meeting notes into executive-ready one page summaries with traffic lights, deltas, and blockers. Designed to run entirely on GitHub Pages – clone, push, and enable Pages to ship.

## ✨ Features

- **Smartsheet integration** – paste an access token and sheet ID to pull rows client-side.
- **Google Docs notes** – import shared docs (exported as plain text) or paste notes manually.
- **Tone controls** – choose executive, team, detailed, or craft a custom narrative prompt.
- **Traffic lights & deltas** – configurable column mapping for status, week-over-week movement, and blockers.
- **Exports** – download the generated one pager as PDF, PowerPoint, or DOCX.
- **Improvement ideas** – baked-in suggestions for growing the product beyond MVP.

## 🚀 Getting started

1. Clone the repository and push it to your GitHub account.
2. Enable GitHub Pages for the repository (root directory). The static assets require no build step.
3. Browse to the published site.

## 🔌 Connecting data sources

### Smartsheet

1. Generate a Smartsheet API access token with read permissions.
2. Enter the token and sheet ID, then map the relevant column titles (status, delta, blockers, owner).
3. (Optional) If Smartsheet rejects the browser request with “Failed to fetch”, add a proxy template URL that includes a `{target}` placeholder. The app replaces the placeholder with the Smartsheet API endpoint.
4. Click **Load sheet** – data stays in the browser only.

#### Setting up a personal CORS proxy

Smartsheet's public API does not return the CORS headers required for direct browser calls. When this happens the app will surface a message suggesting a proxy. The quickest approach is to deploy a free [Cloudflare Worker](https://developers.cloudflare.com/workers/) and point the **Optional CORS proxy template** field to it (e.g. `https://your-worker.workers.dev/?target={target}`).

```js
export default {
  async fetch(request) {
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        status: 204,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Methods': 'GET,OPTIONS',
          'Access-Control-Allow-Headers': request.headers.get('Access-Control-Request-Headers') ?? 'authorization,accept',
        },
      });
    }

    const url = new URL(request.url);
    const target = url.searchParams.get('target');
    if (!target) {
      return new Response('Missing target parameter', { status: 400 });
    }

    const upstream = await fetch(target, {
      headers: {
        Authorization: request.headers.get('Authorization') ?? '',
        Accept: 'application/json',
      },
    });

    return new Response(upstream.body, {
      status: upstream.status,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Content-Type': upstream.headers.get('Content-Type') ?? 'application/json',
      },
    });
  },
};
```

Publish the worker and paste its URL (with `{target}`) into the Smartsheet panel. The access token remains in the Authorization header and is only transmitted to your worker and Smartsheet.

### Google Docs

1. Share your meeting notes document with “Anyone with the link – Viewer”.
2. Paste the document ID (the long string in the URL) and click **Load Google Doc**.
3. Alternatively, paste notes directly into the textarea.

## 🧠 Generating the report

1. Choose the tone for the narrative and optionally add emphasis prompts.
2. Press **Generate Report** to build the one page summary.
3. Export via PDF, PPTX, or DOCX buttons.

## 🛡️ Security

- No credentials are persisted or sent to a backend; everything runs within your browser session.
- Revoke API tokens when you are done.
- When using a custom proxy, host it under your control. Third-party proxies can log your Smartsheet access token.

## 💡 Future enhancements

- OAuth-powered Google and Microsoft connectors for secure enterprise deployment.
- Historical Smartsheet snapshots to quantify velocity and progress trends automatically.
- AI-generated mitigations and owner nudges based on blocker patterns.
- Slack/Teams publishing workflows with tone-aware variants per stakeholder group.

## 🛠️ Development

Static site – edit `index.html`, `assets/style.css`, and `assets/app.js`. No build tooling required.
