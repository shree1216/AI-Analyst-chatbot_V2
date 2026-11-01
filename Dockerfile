FROM python:3.11-slim

   # Microsoft ODBC Driver 18 is required to reach SQL Server from Linux; unixodbc alone is not enough.
   RUN apt-get update && apt-get install -y --no-install-recommends curl ca-certificates unixodbc \
       && . /etc/os-release \
       && curl -fsSL -o /tmp/packages-microsoft-prod.deb "https://packages.microsoft.com/config/debian/${VERSION_ID}/packages-microsoft-prod.deb" \
       && dpkg -i /tmp/packages-microsoft-prod.deb && rm /tmp/packages-microsoft-prod.deb \
       && apt-get update && ACCEPT_EULA=Y apt-get install -y --no-install-recommends msodbcsql18 \
       && rm -rf /var/lib/apt/lists/*

   WORKDIR /app

   COPY requirements.txt .
   RUN pip install --no-cache-dir -r requirements.txt

   COPY . .

   EXPOSE 8501
   CMD streamlit run employee_sql_chatbot.py --server.port=${PORT:-8501} --server.address=0.0.0.0
