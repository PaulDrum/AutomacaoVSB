import random

def gerar_numeros():
    return random.sample(range(1, 61), 6)

def main():
    for _ in range(2):
        numeros_aleatorios = gerar_numeros()
        print("Par de 6 números:", numeros_aleatorios)

if __name__ == "__main__":
    main()
