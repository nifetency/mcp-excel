FROM python:3.12-slim

WORKDIR /app

# Install deps first (cached layer unless requirements change)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App code
COPY mcp_excel_server.py .

# creds.json is mounted at runtime -- never baked into the image
ENV GOOGLE_APPLICATION_CREDENTIALS=/secrets/creds.json
ENV MCP_TRANSPORT=http
ENV MCP_PORT=5002

EXPOSE 5002

CMD ["python", "mcp_excel_server.py"]
