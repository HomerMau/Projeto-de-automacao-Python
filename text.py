# Importando as bibliotecas necessárias
import matplotlib.pyplot as plt
import numpy as np

# Definindo os dados da tabela como listas
tempo = [4, 8, 12, 16, 20, 24]
densidade = [0.02, 0.04, 0.08, 0.16, 0.32, 0.64]

# Plotando o gráfico de dispersão dos dados
plt.scatter(tempo, densidade, color='blue', label='Dados experimentais')

# Ajustando uma curva polinomial de grau 2 aos dados
coef = np.polyfit(tempo, densidade, 2)
poli = np.poly1d(coef)

# Plotando a curva ajustada aos dados
x = np.linspace(4, 24, 100)
y = poli(x)
plt.plot(x, y, color='red', label='Curva ajustada')

# Adicionando título e legendas ao gráfico
plt.title('Crescimento de Corynebacterium pseudotuberculosis')
plt.xlabel('Tempo (h)')
plt.ylabel('Densidade óptica')
plt.legend()

# Mostrando o gráfico na tela
plt.show()
