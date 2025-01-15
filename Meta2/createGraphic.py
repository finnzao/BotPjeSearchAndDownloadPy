import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# Carregar o DataFrame atualizado
df_resultado = pd.read_excel('original_atualizado.xlsx')

# Preencher valores nulos
df_resultado['fila'] = df_resultado['fila'].fillna('Desconhecido')
df_resultado['orgaoJulgador'] = df_resultado['orgaoJulgador'].fillna('Desconhecido')

# Agrupar dados e calcular quantidades e porcentagens
df_counts = df_resultado.groupby(['fila', 'orgaoJulgador']).size().reset_index(name='quantidade')
total_processos = df_counts['quantidade'].sum()
df_counts['porcentagem'] = (df_counts['quantidade'] / total_processos) * 100

# Preparar dados para o gráfico Sunburst
# Nós da 'fila'
df_fila_total = df_resultado.groupby('fila').size().reset_index(name='quantidade')
df_sunburst_fila = pd.DataFrame({
    'id': df_fila_total['fila'],
    'label': df_fila_total['fila'],
    'parent': [''] * len(df_fila_total),
    'value': df_fila_total['quantidade']
})

# Nós do 'orgaoJulgador'
df_sunburst_orgao = pd.DataFrame({
    'id': df_counts['fila'] + '/' + df_counts['orgaoJulgador'],
    'label': df_counts['orgaoJulgador'],
    'parent': df_counts['fila'],
    'value': df_counts['quantidade']
})

# Combinar dados para o Sunburst
df_sunburst = pd.concat([df_sunburst_fila, df_sunburst_orgao], ignore_index=True)

# Criar figura com subplots
fig = make_subplots(
    rows=1, cols=2,
    specs=[[{'type': 'domain'}, {'type': 'table'}]],
    column_widths=[0.6, 0.4]
)

# Adicionar gráfico Sunburst
sunburst_trace = go.Sunburst(
    ids=df_sunburst['id'],
    labels=df_sunburst['label'],
    parents=df_sunburst['parent'],
    values=df_sunburst['value'],
    branchvalues='total',
    hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Porcentagem: %{percentParent:.2%}<extra></extra>'
)
fig.add_trace(sunburst_trace, row=1, col=1)

# Adicionar tabela
table_trace = go.Table(
    header=dict(
        values=['<b>Fila</b>', '<b>Órgão Julgador</b>', '<b>Quantidade</b>', '<b>Porcentagem (%)</b>'],
        fill_color='paleturquoise',
        align='left'
    ),
    cells=dict(
        values=[
            df_counts['fila'],
            df_counts['orgaoJulgador'],
            df_counts['quantidade'],
            df_counts['porcentagem'].round(2)
        ],
        fill_color='lavender',
        align='left'
    )
)
fig.add_trace(table_trace, row=1, col=2)

# Atualizar layout
fig.update_layout(
    title_text='Distribuição de Processos por Fila e Órgão Julgador',
    showlegend=False
)

# Exibir figura
fig.show()
