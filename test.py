import os

def list_files_with_labels(root_directory):
    file_labels = {}
    
    for dirpath, dirnames, filenames in os.walk(root_directory):
        # Determina o rótulo com base no diretório
        if dirpath == root_directory:
            label = os.path.basename(root_directory)
        else:
            label = os.path.basename(dirpath)
        
        # Atribui o rótulo a cada arquivo no diretório
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            if label not in file_labels:
                file_labels[label] = []
            file_labels[label].append(file_path)
    
    return file_labels

# Exemplo de uso
root_directory = 'C:/Users/s-Lucas.SAraujo/Códigos'
file_labels = list_files_with_labels(root_directory)

# Imprime os arquivos com seus rótulos
for label, files in file_labels.items():
    print(f"Rótulo: {label}")
    for file in files:
        print(f"  {file}")