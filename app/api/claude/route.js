// This runs on the server (Vercel serverless function), never in the browser.
// The API key lives here, in an environment variable — it's never sent to the client.
//
// Set ANTHROPIC_API_KEY in your Vercel project settings (or .env.local for local dev).
// Get a key at https://console.anthropic.com

export async function POST(request) {
  const apiKey = process.env.ANTHROPIC_API_KEY;

  if (!apiKey) {
    return Response.json(
      { error: "ANTHROPIC_API_KEY is not set on the server. Add it in Vercel → Project Settings → Environment Variables, then redeploy." },
      { status: 500 }
    );
  }

  let body;
  try {
    body = await request.json();
  } catch (err) {
    return Response.json({ error: "Invalid request body" }, { status: 400 });
  }

  const { messages, maxTokens } = body || {};
  if (!Array.isArray(messages) || messages.length === 0) {
    return Response.json({ error: "Missing 'messages' array" }, { status: 400 });
  }

  try {
    const anthropicResponse = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: maxTokens || 1000,
        messages,
      }),
    });

    const data = await anthropicResponse.json();

    // Pass Anthropic's own status code straight through (429, 529, etc.) so the client's
    // existing retry/backoff logic keeps working exactly as it did calling Anthropic directly.
    return Response.json(data, { status: anthropicResponse.status });
  } catch (err) {
    return Response.json({ error: `Failed to reach Anthropic API: ${err.message}` }, { status: 502 });
  }
}
