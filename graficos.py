import pandas as pd
import matplotlib.pyplot as plt

# Defina o caminho do arquivo
caminho_arquivo = r'C:\pythonjr\RelatScripts\dados_hibrida1.xlsx'

# Leia o arquivo Excel especificando a matriz desejada
df_trafo_antes = pd.read_excel(caminho_arquivo, sheet_name=0, usecols='B:G', skiprows=2, nrows=145)
df_trafo_depois = pd.read_excel(caminho_arquivo, sheet_name=0, usecols='P:U', skiprows=2, nrows=145)
df_red_antes = pd.read_excel(caminho_arquivo, sheet_name=0, usecols='H:M', skiprows=2, nrows=145)
df_red_depois = pd.read_excel(caminho_arquivo, sheet_name=0, usecols='V:AA', skiprows=2, nrows=145)
df_limites = pd.read_excel(caminho_arquivo, sheet_name=0, usecols='AC:AE', skiprows=2, nrows=145)

# Função para converter a primeira coluna para o formato hh:mm
def converter_para_horas(dataframe):
    dataframe.iloc[:, 0] = pd.to_datetime(dataframe.iloc[:, 0], format='%H:%M:%S', errors='coerce').dt.strftime('%H:%M')
    return dataframe

# Converter as primeiras colunas de cada DataFrame
df_trafo_antes = converter_para_horas(df_trafo_antes)
df_trafo_depois = converter_para_horas(df_trafo_depois)
df_red_antes = converter_para_horas(df_red_antes)
df_red_depois = converter_para_horas(df_red_depois)

# Criação da figura com subplots
fig1, axs = plt.subplots(1, 2, figsize=(12, 6), sharey=True)

# Primeiro gráfico baseado no DataFrame df_trafo_antes
axs[0].step(df_trafo_antes.iloc[:, 0], df_trafo_antes.iloc[:, 1], label='Fase A', color='red')
axs[0].step(df_trafo_antes.iloc[:, 0], df_trafo_antes.iloc[:, 2], label='Fase B', color='blue')
axs[0].step(df_trafo_antes.iloc[:, 0], df_trafo_antes.iloc[:, 3], label='Fase C', color='green')
axs[0].axhspan(203, 230, color='lightgreen', alpha=0.3)  # Faixa de tensão
axs[0].set_xlabel('Horário', fontsize=15)
axs[0].set_ylabel('Tensão (V)', fontsize=15)
axs[0].set_title('TENSÃO NO TRANSFORMADOR (ANTES)')
axs[0].set_ylim(190, 240)
axs[0].legend(fontsize=11, loc='upper right')
axs[0].tick_params(axis='x', labelsize=10, rotation=90)
axs[0].tick_params(axis='y', labelsize=12)
axs[0].grid(True, linestyle=':', color='lightgray')
axs[0].set_yticks(range(190, 241, 5))

# Ajuste dos rótulos do eixo X para mostrar apenas algumas horas para evitar sobreposição
num_labels = 30  # Número desejado de rótulos no eixo X
ticks_to_use = df_trafo_antes.iloc[::len(df_trafo_antes)//num_labels, 0]
axs[0].set_xticks(ticks_to_use.index)
axs[0].set_xticklabels(ticks_to_use, rotation=90)

# Segundo gráfico baseado no DataFrame df_trafo_depois
axs[1].step(df_trafo_depois.iloc[:, 0], df_trafo_depois.iloc[:, 1], label='Fase A', color='red')
axs[1].step(df_trafo_depois.iloc[:, 0], df_trafo_depois.iloc[:, 2], label='Fase B', color='blue')
axs[1].step(df_trafo_depois.iloc[:, 0], df_trafo_depois.iloc[:, 3], label='Fase C', color='green')
axs[1].axhspan(203, 230, color='lightgreen', alpha=0.3)  # Faixa de tensão
axs[1].set_xlabel('Horário', fontsize=15)
axs[1].set_title('TENSÃO NO TRANSFORMADOR (DEPOIS)')
axs[1].set_ylim(190, 240)
axs[1].legend(fontsize=11, loc='upper right')
axs[1].tick_params(axis='x', labelsize=10, rotation=90)
axs[1].tick_params(axis='y', labelsize=12)
axs[1].grid(True, linestyle=':', color='lightgray')
axs[1].set_yticks(range(190, 241, 5))

# Ajuste dos rótulos do eixo X para mostrar apenas algumas horas para evitar sobreposição
ticks_to_use = df_trafo_depois.iloc[::len(df_trafo_depois)//num_labels, 0]
axs[1].set_xticks(ticks_to_use.index)
axs[1].set_xticklabels(ticks_to_use, rotation=90)

plt.tight_layout()  # Ajusta automaticamente o espaçamento do subplot para evitar sobreposição
plt.show()

# Criação da Figura 2 com subplots para df_red_antes e df_red_depois
fig2, axs2 = plt.subplots(1, 2, figsize=(12, 6), sharey=True)

# Primeiro gráfico baseado no DataFrame df_red_antes
axs2[0].step(df_trafo_antes.iloc[:, 0], df_red_antes.iloc[:, 1], label='Fase A', color='red')
axs2[0].step(df_trafo_antes.iloc[:, 0], df_red_antes.iloc[:, 2], label='Fase B', color='blue')
axs2[0].step(df_trafo_antes.iloc[:, 0], df_red_antes.iloc[:, 4], label='Fase C', color='green')
axs2[0].axhspan(203, 230, color='lightgreen', alpha=0.3)  # Faixa de tensão
axs2[0].set_xlabel('Horário', fontsize=15)
axs2[0].set_ylabel('Tensão (V)', fontsize=15)
axs2[0].set_title('TENSÃO NO RED (ANTES)')
axs2[0].set_ylim(190, 240)
axs2[0].legend(fontsize=11, loc='upper right')
axs2[0].tick_params(axis='x', labelsize=10, rotation=90)
axs2[0].tick_params(axis='y', labelsize=12)
axs2[0].grid(True, linestyle=':', color='lightgray')
axs2[0].set_yticks(range(190, 241, 5))

# Ajuste dos rótulos do eixo X para mostrar apenas algumas horas para evitar sobreposição
num_labels = 30  # Número desejado de rótulos no eixo X
ticks_to_use = df_trafo_antes.iloc[::len(df_red_antes)//num_labels, 0]
axs2[0].set_xticks(ticks_to_use.index)
axs2[0].set_xticklabels(ticks_to_use, rotation=90)

# Segundo gráfico baseado no DataFrame df_red_depois
axs2[1].step(df_trafo_depois.iloc[:, 0], df_red_depois.iloc[:, 1], label='Fase A', color='red')
axs2[1].step(df_trafo_depois.iloc[:, 0], df_red_depois.iloc[:, 2], label='Fase B', color='blue')
axs2[1].step(df_trafo_depois.iloc[:, 0], df_red_depois.iloc[:, 3], label='Fase C', color='green')
axs2[1].axhspan(203, 230, color='lightgreen', alpha=0.3)  # Faixa de tensão
axs2[1].set_xlabel('Horário', fontsize=15)
axs2[1].set_title('TENSÃO NO RED (DEPOIS)')
axs2[1].set_ylim(190, 240)
axs2[1].legend(fontsize=11, loc='upper right')
axs2[1].tick_params(axis='x', labelsize=10, rotation=90)
axs2[1].tick_params(axis='y', labelsize=12)
axs2[1].grid(True, linestyle=':', color='lightgray')
axs2[1].set_yticks(range(190, 241, 5))

# Ajuste dos rótulos do eixo X para mostrar apenas algumas horas para evitar sobreposição
ticks_to_use = df_trafo_depois.iloc[::len(df_red_depois)//num_labels, 0]
axs2[1].set_xticks(ticks_to_use.index)
axs2[1].set_xticklabels(ticks_to_use, rotation=90)

plt.tight_layout()  # Ajusta automaticamente o espaçamento do subplot para evitar sobreposição
plt.show()


# Criação da figura com subplots
fig3, axs3 = plt.subplots(1, 2, figsize=(12, 6), sharey=True)

# Primeiro gráfico baseado no DataFrame df_trafo_antes
axs3[0].step(df_trafo_antes.iloc[:, 0], df_trafo_antes.iloc[:, 5], label='kVA', color='red')
axs3[0].axhspan(0, 135, color='lightgreen', alpha=0.3)  # Faixa de tensão
axs3[0].set_xlabel('Horário', fontsize=15)
axs3[0].set_ylabel('Potência Aparente (kVA)', fontsize=15)
axs3[0].set_title('CARREGAMENTO DO TRANSFORMADOR (ANTES)')
axs3[0].set_ylim(0, 150)
axs3[0].legend(fontsize=11, loc='upper right')
axs3[0].tick_params(axis='x', labelsize=10, rotation=90)
axs3[0].tick_params(axis='y', labelsize=12)
axs3[0].grid(True, linestyle=':', color='lightgray')
axs3[0].set_yticks(range(0, 150, 10))

# Ajuste dos rótulos do eixo X para mostrar apenas algumas horas para evitar sobreposição
num_labels = 30  # Número desejado de rótulos no eixo X
ticks_to_use = df_trafo_antes.iloc[::len(df_trafo_antes)//num_labels, 0]
axs3[0].set_xticks(ticks_to_use.index)
axs3[0].set_xticklabels(ticks_to_use, rotation=90)

# Segundo gráfico baseado no DataFrame df_trafo_depois
axs3[1].step(df_trafo_depois.iloc[:, 0], df_trafo_depois.iloc[:, 5], label='kVA', color='red')
axs3[1].axhspan(0, 135, color='lightgreen', alpha=0.3)  # Faixa de tensão
axs3[1].set_xlabel('Horário', fontsize=15)
axs3[1].set_title('CARREGAMENTO DO TRANSFORMADOR (DEPOIS)')
axs3[1].set_ylim(0, 150)
axs3[1].legend(fontsize=11, loc='upper right')
axs3[1].tick_params(axis='x', labelsize=10, rotation=90)
axs3[1].tick_params(axis='y', labelsize=12)
axs3[1].grid(True, linestyle=':', color='lightgray')
axs3[1].set_yticks(range(0, 150, 10))

# Ajuste dos rótulos do eixo X para mostrar apenas algumas horas para evitar sobreposição
ticks_to_use = df_trafo_depois.iloc[::len(df_trafo_depois)//num_labels, 0]
axs3[1].set_xticks(ticks_to_use.index)
axs3[1].set_xticklabels(ticks_to_use, rotation=90)

plt.tight_layout()  # Ajusta automaticamente o espaçamento do subplot para evitar sobreposição
plt.show()


# Criação da figura com subplots
fig4, axs4 = plt.subplots(1, 2, figsize=(12, 6), sharey=True)

# Primeiro gráfico baseado no DataFrame df_trafo_antes
axs4[0].step(df_trafo_antes.iloc[:, 0], df_trafo_antes.iloc[:, 4], label='kW', color='red')
axs4[0].axhspan(-100, 0, color='lightgrAY', alpha=0.3)  # Faixa de tensão
axs4[0].set_xlabel('Horário', fontsize=15)
axs4[0].set_ylabel('Potência Aparente (kVA)', fontsize=15)
axs4[0].set_title('POTÊNCIA ATIVA NO TRANSFORMADOR (ANTES)')
axs4[0].set_ylim(-100, 50)
axs4[0].legend(fontsize=11, loc='upper right')
axs4[0].tick_params(axis='x', labelsize=10, rotation=90)
axs4[0].tick_params(axis='y', labelsize=12)
axs4[0].grid(True, linestyle=':', color='lightgray')
axs4[0].set_yticks(range(-100, 50, 10))

# Ajuste dos rótulos do eixo X para mostrar apenas algumas horas para evitar sobreposição
num_labels = 30  # Número desejado de rótulos no eixo X
ticks_to_use = df_trafo_antes.iloc[::len(df_trafo_antes)//num_labels, 0]
axs4[0].set_xticks(ticks_to_use.index)
axs4[0].set_xticklabels(ticks_to_use, rotation=90)

# Segundo gráfico baseado no DataFrame df_trafo_depois
axs4[1].step(df_trafo_depois.iloc[:, 0], df_trafo_depois.iloc[:, 4], label='kW', color='red')
axs4[1].axhspan(-100, 0, color='lightgrAY', alpha=0.3)  # Faixa de tensão
axs4[1].set_xlabel('Horário', fontsize=15)
axs4[1].set_title('POTÊNCIA ATIVA NO TRANSFORMADOR (DEPOIS)')
axs4[1].set_ylim(-100, 50)
axs4[1].legend(fontsize=11, loc='upper right')
axs4[1].tick_params(axis='x', labelsize=10, rotation=90)
axs4[1].tick_params(axis='y', labelsize=12)
axs4[1].grid(True, linestyle=':', color='lightgray')
axs4[1].set_yticks(range(-100, 50, 10))

# Ajuste dos rótulos do eixo X para mostrar apenas algumas horas para evitar sobreposição
ticks_to_use = df_trafo_depois.iloc[::len(df_trafo_depois)//num_labels, 0]
axs4[1].set_xticks(ticks_to_use.index)
axs4[1].set_xticklabels(ticks_to_use, rotation=90)

plt.tight_layout()  # Ajusta automaticamente o espaçamento do subplot para evitar sobreposição
plt.show()