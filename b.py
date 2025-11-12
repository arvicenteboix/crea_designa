from google import genai

# The client gets the API key from the environment variable `GEMINI_API_KEY`.
client = genai.Client(api_key="AIzaSyDGQwD4hLNFNLl51LIqLdG8U6h8t6DgsgU")

response = client.models.generate_content(
    model="gemini-2.5-flash", contents="Convertix en valencià el número 378 a lletres. Dona'm només la resposta sense cap explicació addicional."
)
print(response.text)