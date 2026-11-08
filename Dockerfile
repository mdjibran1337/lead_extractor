FROM python:3.12-slim

# 1. System dependencies install karein (Chrome aur keyring ke liye)
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    unzip \
    curl \
    && rm -rf /var/lib/apt/lists/*

# 2. Google Chrome ki GPG key ko naye tarike se add karein (apt-key fix)
RUN curl -fsSL https://dl.google.com/linux/linux_signing_key.pub \
    | gpg --dearmor -o /usr/share/keyrings/google-chrome.gpg

# 3. Repository add karein aur Chrome install karein
RUN echo "deb [arch=amd64 signed-by=/usr/share/keyrings/google-chrome.gpg] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list \
    && apt-get update \
    && apt-get install -y google-chrome-stable \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Port 3000 par server chalayein
CMD ["uvicorn", "app.py:app", "--host", "0.0.0.0", "--port", "3000"]
