# Excel MCP Server - Deployment

## Quick Start

```bash
cd deploy
docker compose up -d
```

## Services

| Service | Port | URL |
|---------|-----|-----|
| MCP Server | 3749 | http://localhost:3749 |
| Frontend | 3750 | http://localhost:3750 |

## Endpoints

- `GET /health` - Health check
- `GET /history` - Chat history
- `POST /chat` - Chat with AI
- `POST /clear-history` - Clear history

## Environment Variables

- `PORT` - Server port (default: 3749)
- `OPENCODE_MODEL` - AI model (default: opencode)
- `MAX_HISTORY` - Max history items (default: 50)
