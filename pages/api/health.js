/**
 * GET /api/health
 * Simple health check. Returns 200 if the service is up and the API key is configured.
 */
export default function handler(req, res) {
  const hasKey = !!process.env.ANTHROPIC_API_KEY;
  return res.status(200).json({
    status: 'ok',
    anthropicKeyConfigured: hasKey,
    timestamp: new Date().toISOString(),
  });
}
