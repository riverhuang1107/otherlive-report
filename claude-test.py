import anthropic
import os

# The Anthropic Python SDK sends the anthropic-version header automatically.
# Project scoping uses the anthropic-workspace-id header.
client = anthropic.Anthropic(
    default_headers={"anthropic-workspace-id": os.environ["ANTHROPIC_WORKSPACE_ID"]},
)

message = client.messages.create(
    #model="anthropic.claude-haiku-4-5",
    model="anthropic.claude-haiku-4-5-20251001-v1:0",
    max_tokens=64,
    messages=[{"role": "user", "content": "Amazon Bedrock 是什么？"}],
)
print(message.content[0].text)
