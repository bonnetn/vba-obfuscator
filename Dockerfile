FROM python:3.7-alpine

WORKDIR /obfuscator

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD [ "python", "./obfuscate.py" ]
