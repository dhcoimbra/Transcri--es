import requests
import time

# Coloque aqui a sua chave de API AssemblyAI
ASSEMBLYAI_API_KEY = "ddc0694917fb46ee9588b692540abc28"

# Função para fazer o upload do arquivo de áudio para a AssemblyAI
def upload_audio(audio_path):
    headers = {
        'authorization': ASSEMBLYAI_API_KEY,
    }

    with open(audio_path, 'rb') as f:
        audio_data = f.read()

    response = requests.post('https://api.assemblyai.com/v2/upload', headers=headers, data=audio_data)

    if response.status_code == 200:
        return response.json()['upload_url']
    else:
        return None

# Função para transcrever o áudio usando a API da AssemblyAI
def transcrever_audio_assemblyai(audio_url):
    headers = {
        'authorization': ASSEMBLYAI_API_KEY,
        'content-type': 'application/json'
    }

    transcript_request = {
        'audio_url': audio_url,
        'language_code': 'pt'  # Define o idioma como português
    }

    response = requests.post('https://api.assemblyai.com/v2/transcript', json=transcript_request, headers=headers)

    if response.status_code == 200:
        transcript_id = response.json()['id']
    else:
        return None

    # Aguardar a conclusão da transcrição
    while True:
        status_response = requests.get(f'https://api.assemblyai.com/v2/transcript/{transcript_id}', headers=headers)
        status = status_response.json()

        if status['status'] == 'completed':
            return status['text']
        elif status['status'] == 'failed':
            return None
        else:
            time.sleep(5)
