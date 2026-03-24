# Uniram Support Jennifer — AI Email Auto-Reply System

Monitors `support@uniram.com` and automatically handles incoming support emails using GPT-4o-mini and a ChromaDB knowledge base.

## Architecture

```
support@uniram.com (Inbox)
    ↓ Azure Function Timer Trigger (every 30 min)
    ↓ GPT classification
    ├── Technical question (confidence ≥ 0.65) → Jennifer replies directly
    ├── Technical question (confidence < 0.65) → Escalate to ken@uniram.com
    ├── Vague / no details → Jennifer asks for more info
    ├── Pricing inquiry → Forward to sales@uniram.com (with original email + images)
    ├── Safety-critical (solvents/chemicals) → Force escalate to Ken
    └── Spam / auto-reply → Skip silently
    
    Processed emails → moved to "Processed" subfolder
```

## Learning

- **From Ken's feedback**: Scans Ken's inbox for replies to `[Support Escalation]` emails. Uses GPT to parse intent (junk / handled / sales lead) and updates the knowledge base accordingly.
- **From historical folders**: One-time scan of 2023/2024/2025/Ken/Finn folders to extract Q&A pairs.

## Azure Resources

| Resource | Name |
|----------|------|
| Function App | `uniram-support` |
| Resource Group | `uniram-reports-rg` |
| Storage Account | `uniramreportsrg98d2` |
| KB Container | `uniram-support-kb` |

## CI/CD

Push to `main` → GitHub Actions → auto-deploy to Azure Function App.

**Required GitHub Secret:** `AZURE_CREDENTIALS` (Azure Service Principal JSON)

## Environment Variables

| Variable | Description |
|----------|-------------|
| `GRAPH_TENANT_ID` | Azure AD Tenant ID |
| `GRAPH_CLIENT_ID` | App Registration Client ID |
| `GRAPH_CLIENT_SECRET` | App Registration Secret |
| `OPENAI_API_KEY` | OpenAI API Key (Uniram Tech Support Agent) |
| `SUPPORT_MAILBOX` | `support@uniram.com` |
| `AzureWebJobsStorage` | Azure Storage connection string (for KB persistence) |
| `LEARN_HISTORY` | `true` to enable historical folder learning |
| `LEARN_HISTORY_FOLDER` | Current folder being processed (auto-advances) |

Last deployed: 2026-03-24 02:05 UTC
