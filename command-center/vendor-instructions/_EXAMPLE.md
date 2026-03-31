# Vendor Instructions Template

Save a copy of this file as `{vendor-slug}.md` (e.g., `acme-corp.md`) to add
per-vendor instructions that the AI will follow when processing this vendor.

These instructions are loaded automatically before every analysis and override
default behavior. Use them to:

- Correct misinterpretations ("When they say X, they actually mean Y")
- Set vendor-specific rules ("Never offer discounts to this vendor")
- Record context the AI keeps getting wrong ("Their billing contact is Jane, not John")
- Note special handling ("Always CC legal@company.com on contract emails")
- Track quirks ("They send from noreply@ but replies go to support@")

## Example

```
- When this vendor mentions "net terms", they mean NET-60, not NET-30.
- Their account manager changed from Sarah to Mike in Jan 2026. Always address Mike.
- Do NOT auto-archive their emails about "monthly reconciliation" - these need manual review.
- They use "pause" to mean temporary hold, not cancellation. Don't update status to dead.
- Always check their Box folder for the latest rate card before responding to pricing questions.
```
