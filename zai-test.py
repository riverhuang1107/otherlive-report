from openai import OpenAI

client = OpenAI()

response = client.chat.completions.create(
    model="zai.glm-5",
    messages=[{"role": "user", "content": "Can you explain the features of Amazon Bedrock?"}]
    )
print(response)
