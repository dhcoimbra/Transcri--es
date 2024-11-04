import os
from datasets import Dataset, DatasetDict
from PIL import Image
from transformers import ViTFeatureExtractor, ViTForImageClassification, Trainer, TrainingArguments
import torch

# Caminho para o diretório de imagens, com subpastas 'comprovante' e 'outro'
base_dir = r"C:\modelos\imagens"

# Inicializar o feature extractor para ViT
feature_extractor = ViTFeatureExtractor.from_pretrained("google/vit-base-patch16-224-in21k")

# Função para carregar imagens e atribuir rótulos com base no diretório
def load_images_from_directory(base_dir):
    data = {"image_path": [], "label": []}
    class_to_label = {"comprovante": 1, "outro": 0}  # Definindo rótulos: 1 para comprovante, 0 para outro

    # Percorre cada pasta (classe) e rotula as imagens
    for class_name, label in class_to_label.items():
        class_dir = os.path.join(base_dir, class_name)
        for image_name in os.listdir(class_dir):
            image_path = os.path.join(class_dir, image_name)
            data["image_path"].append(image_path)
            data["label"].append(label)

    # Retorna o dataset carregado com as imagens e rótulos
    return Dataset.from_dict(data)

# Carregar as imagens e criar o dataset
raw_dataset = load_images_from_directory(base_dir)

# Função de pré-processamento para ViT
def preprocess_data(example):
    # Carregar imagem e converter para RGB
    image = Image.open(example['image_path']).convert("RGB")
    # Aplicar o feature extractor para obter o tensor de imagem
    encoding = feature_extractor(images=image, return_tensors="pt")
    
    # Squeeze apenas nos tensores de imagem
    encoding = {key: val.squeeze() for key, val in encoding.items()}
    
    # Adicionar o rótulo separadamente
    encoding["label"] = example["label"]
    return encoding

# Aplicar o pré-processamento no dataset e remover colunas desnecessárias
processed_dataset = raw_dataset.map(preprocess_data, remove_columns=["image_path"])

# Dividir o dataset em treino e validação
dataset = processed_dataset.train_test_split(test_size=0.2)
dataset_dict = DatasetDict({"train": dataset["train"], "validation": dataset["test"]})

# Configurar o treinamento
training_args = TrainingArguments(
    output_dir=r'C:\modelos\imagens\vit_comprovante',
    evaluation_strategy="epoch",
    per_device_train_batch_size=4,
    per_device_eval_batch_size=4,
    num_train_epochs=3,
    save_strategy="epoch",
    load_best_model_at_end=True,
)

# Carregar o modelo ViT para classificação
model = ViTForImageClassification.from_pretrained("google/vit-base-patch16-224-in21k", num_labels=2)

# Configurar o treinador
trainer = Trainer(
    model=model,
    args=training_args,
    train_dataset=dataset_dict['train'],
    eval_dataset=dataset_dict['validation'],
)

# Iniciar o treinamento
trainer.train()

# Salvar o modelo e o feature extractor treinados
model.save_pretrained(r'C:\modelos\imagens\vit_comprovante')
feature_extractor.save_pretrained(r'C:\modelos\imagens\vit_comprovante')

print("Treinamento concluído e modelo salvo em './vit_comprovante'")
