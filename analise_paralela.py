#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Análise Paralela de Colaboradores
=================================
Este script executa análises em paralelo dos arquivos Excel e identifica melhorias
na qualidade da análise baseada na coluna "SITUAÇÃO" de cada colaborador.
"""

import os
import pandas as pd
import numpy as np
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import seaborn as sns
from collections import defaultdict, Counter

# Importações locais
from debug_excel import AnalisadorExcel
from analise_360 import Analise360
from data_analysis_pipeline import DataAnalysisPipeline

def analisar_situacao_colaborador(nome_arquivo, nome_aba):
    """
    Analisa a qualidade dos registros na coluna SITUAÇÃO para um colaborador específico.
    
    Args:
        nome_arquivo (str): Caminho para o arquivo Excel
        nome_aba (str): Nome da aba/colaborador a ser analisada
        
    Returns:
        dict: Dicionário com métricas de qualidade dos registros
    """
    try:
        # Carregar dados do colaborador
        df = pd.read_excel(nome_arquivo, sheet_name=nome_aba)
        
        # Normalizar nomes das colunas
        colunas_normalizadas = []
        for col in df.columns:
            col_norm = str(col).strip().upper()
            if col_norm in ['SITUAÇÂO', 'SITUAÇÃO']:
                col_norm = 'SITUACAO'
            colunas_normalizadas.append(col_norm)
        
        df.columns = colunas_normalizadas
        
        # Verificar se a coluna SITUACAO existe
        if 'SITUACAO' not in df.columns:
            return {
                'colaborador': nome_aba,
                'arquivo': nome_arquivo,
                'erro': 'Coluna SITUACAO não encontrada',
                'status': 'FALHA'
            }
        
        # Análise de qualidade da coluna SITUACAO
        total_registros = len(df)
        registros_vazios = df['SITUACAO'].isna().sum()
        valores_unicos = df['SITUACAO'].dropna().unique()
        contagem_valores = df['SITUACAO'].value_counts().to_dict()
        
        # Verificar padrões de preenchimento
        valores_padronizados = ['PENDENTE', 'VERIFICADO', 'APROVADO', 'QUITADO', 'CANCELADO', 'EM ANÁLISE']
        valores_nao_padronizados = [v for v in valores_unicos if v not in valores_padronizados]
        
        # Verificar se há atualizações diárias
        tem_data = False
        atualizacoes_diarias = {}
        coluna_data = None
        
        # Verificar se há coluna de data
        for col in df.columns:
            if 'DATA' in col:
                tem_data = True
                coluna_data = col
                try:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                    # Agrupar por data e contar atualizações de status
                    atualizacoes = df.groupby(df[col].dt.date)['SITUACAO'].count()
                    atualizacoes_diarias = atualizacoes.to_dict()
                except:
                    pass
        
        # Calcular métricas de qualidade
        taxa_preenchimento = (total_registros - registros_vazios) / total_registros if total_registros > 0 else 0
        taxa_padronizacao = (len(valores_unicos) - len(valores_nao_padronizados)) / len(valores_unicos) if len(valores_unicos) > 0 else 0
        
        # Verificar consistência nas atualizações diárias
        consistencia_diaria = 0
        if atualizacoes_diarias:
            # Calcular desvio padrão das atualizações diárias
            std_atualizacoes = np.std(list(atualizacoes_diarias.values()))
            media_atualizacoes = np.mean(list(atualizacoes_diarias.values()))
            
            # Coeficiente de variação (menor é melhor)
            if media_atualizacoes > 0:
                consistencia_diaria = 1 - min(1, std_atualizacoes / media_atualizacoes)
        
        # Calcular score geral de qualidade
        score_qualidade = (
            0.4 * taxa_preenchimento +  # 40% para preenchimento
            0.3 * taxa_padronizacao +   # 30% para padronização
            0.3 * consistencia_diaria   # 30% para consistência diária
        ) * 100
        
        # Análise de transições de estado (se houver coluna de data)
        analise_transicoes = {}
        if tem_data and coluna_data and not df[coluna_data].isna().all():
            # Ordenar por data
            df_ordenado = df.sort_values(by=coluna_data)
            
            # Verificar transições de estado
            if 'SITUACAO' in df_ordenado.columns and len(df_ordenado) > 1:
                transicoes = []
                situacao_anterior = None
                
                for idx, row in df_ordenado.iterrows():
                    situacao_atual = row['SITUACAO']
                    if pd.notna(situacao_anterior) and pd.notna(situacao_atual) and situacao_anterior != situacao_atual:
                        transicoes.append((situacao_anterior, situacao_atual))
                    situacao_anterior = situacao_atual
                
                # Contar transições
                contagem_transicoes = Counter(transicoes)
                analise_transicoes = {f"{de} -> {para}": contagem for (de, para), contagem in contagem_transicoes.items()}
        
        # Análise de tempo médio em cada situação
        tempos_medios = {}
        if tem_data and coluna_data and not df[coluna_data].isna().all():
            # Agrupar por situação e calcular tempo médio
            df_ordenado = df.sort_values(by=coluna_data)
            df_ordenado['data_anterior'] = df_ordenado[coluna_data].shift(1)
            df_ordenado['tempo_no_estado'] = (df_ordenado[coluna_data] - df_ordenado['data_anterior']).dt.days
            
            # Calcular tempo médio por situação
            tempos_por_situacao = df_ordenado.groupby('SITUACAO')['tempo_no_estado'].mean()
            tempos_medios = tempos_por_situacao.to_dict()
        
        # Gerar visualização da distribuição de situações
        grafico_path = None
        try:
            if contagem_valores:
                plt.figure(figsize=(10, 6))
                sns.barplot(x=list(contagem_valores.keys()), y=list(contagem_valores.values()))
                plt.title(f'Distribuição de Situações - {nome_aba}')
                plt.xlabel('Situação')
                plt.ylabel('Quantidade')
                plt.xticks(rotation=45)
                plt.tight_layout()
                
                # Salvar gráfico
                diretorio_graficos = 'graficos_situacao'
                if not os.path.exists(diretorio_graficos):
                    try:
                        os.makedirs(diretorio_graficos)
                    except FileExistsError:
                        # Diretório já existe, podemos continuar
                        pass
                
                # Garantir que o nome do arquivo seja único
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                nome_arquivo_grafico = f'situacao_{nome_aba.replace(" ", "_")}_{timestamp}.png'
                grafico_path = os.path.join(diretorio_graficos, nome_arquivo_grafico)
                plt.savefig(grafico_path)
                plt.close()
        except Exception as e:
            print(f"Erro ao gerar gráfico para {nome_aba}: {str(e)}")
        
        # Identificar problemas e sugestões
        problemas = []
        sugestoes = []
        
        if registros_vazios > 0:
            problemas.append(f"{registros_vazios} registros com SITUACAO vazia")
            sugestoes.append("Preencher todos os campos de SITUACAO")
            
        if valores_nao_padronizados:
            problemas.append(f"{len(valores_nao_padronizados)} valores não padronizados: {', '.join(valores_nao_padronizados)}")
            sugestoes.append("Padronizar valores de SITUACAO para: PENDENTE, VERIFICADO, APROVADO, QUITADO, CANCELADO, EM ANÁLISE")
            
        if not tem_data:
            problemas.append("Não há coluna de DATA para análise temporal")
            sugestoes.append("Adicionar coluna de DATA para permitir análise de atualizações diárias")
            
        if consistencia_diaria < 0.5 and atualizacoes_diarias:
            problemas.append("Baixa consistência nas atualizações diárias")
            sugestoes.append("Manter ritmo constante de atualizações diárias")
        
        return {
            'colaborador': nome_aba,
            'arquivo': nome_arquivo,
            'total_registros': total_registros,
            'registros_vazios': registros_vazios,
            'taxa_preenchimento': taxa_preenchimento * 100,
            'valores_unicos': list(valores_unicos),
            'valores_nao_padronizados': valores_nao_padronizados,
            'taxa_padronizacao': taxa_padronizacao * 100,
            'atualizacoes_diarias': atualizacoes_diarias,
            'consistencia_diaria': consistencia_diaria * 100,
            'score_qualidade': score_qualidade,
            'analise_transicoes': analise_transicoes,
            'tempos_medios': tempos_medios,
            'grafico_path': grafico_path,
            'problemas': problemas,
            'sugestoes': sugestoes,
            'status': 'SUCESSO'
        }
        
    except Exception as e:
        return {
            'colaborador': nome_aba,
            'arquivo': nome_arquivo,
            'erro': str(e),
            'status': 'FALHA'
        }

def analisar_arquivo_paralelo(nome_arquivo):
    """
    Analisa todas as abas de um arquivo Excel em paralelo.
    
    Args:
        nome_arquivo (str): Caminho para o arquivo Excel
        
    Returns:
        dict: Resultados da análise para cada colaborador
    """
    try:
        # Carregar arquivo Excel
        xls = pd.ExcelFile(nome_arquivo)
        
        # Filtrar abas válidas (excluir relatórios gerais, etc.)
        abas_validas = [aba for aba in xls.sheet_names if aba not in ["", "TESTE", "RELATÓRIO GERAL"]]
        
        resultados = {}
        
        # Usar ProcessPoolExecutor para paralelizar a análise
        with ProcessPoolExecutor(max_workers=min(os.cpu_count(), len(abas_validas))) as executor:
            # Criar tarefas para cada aba
            tarefas = {executor.submit(analisar_situacao_colaborador, nome_arquivo, aba): aba for aba in abas_validas}
            
            # Processar resultados conforme são concluídos
            for tarefa in as_completed(tarefas):
                aba = tarefas[tarefa]
                try:
                    resultado = tarefa.result()
                    resultados[aba] = resultado
                except Exception as e:
                    resultados[aba] = {
                        'colaborador': aba,
                        'arquivo': nome_arquivo,
                        'erro': str(e),
                        'status': 'FALHA'
                    }
        
        return resultados
    
    except Exception as e:
        return {'erro_geral': str(e)}

def gerar_relatorio_melhorias(resultados_julio, resultados_leandro):
    """
    Gera um relatório consolidado de melhorias baseado nas análises.
    
    Args:
        resultados_julio (dict): Resultados da análise do grupo Julio
        resultados_leandro (dict): Resultados da análise do grupo Leandro
        
    Returns:
        str: Relatório formatado em texto
    """
    # Combinar resultados
    todos_resultados = {**resultados_julio, **resultados_leandro}
    
    # Calcular estatísticas gerais
    total_colaboradores = len(todos_resultados)
    colaboradores_com_problemas = sum(1 for r in todos_resultados.values() 
                                     if r.get('status') == 'SUCESSO' and r.get('problemas'))
    
    # Agrupar problemas comuns
    todos_problemas = []
    for r in todos_resultados.values():
        if r.get('status') == 'SUCESSO' and r.get('problemas'):
            todos_problemas.extend(r.get('problemas', []))
    
    contagem_problemas = Counter(todos_problemas)
    
    # Agrupar sugestões comuns
    todas_sugestoes = []
    for r in todos_resultados.values():
        if r.get('status') == 'SUCESSO' and r.get('sugestoes'):
            todas_sugestoes.extend(r.get('sugestoes', []))
    
    contagem_sugestoes = Counter(todas_sugestoes)
    
    # Coletar todas as transições de estado
    todas_transicoes = {}
    for r in todos_resultados.values():
        if r.get('status') == 'SUCESSO' and r.get('analise_transicoes'):
            for transicao, contagem in r.get('analise_transicoes', {}).items():
                if transicao in todas_transicoes:
                    todas_transicoes[transicao] += contagem
                else:
                    todas_transicoes[transicao] = contagem
    
    # Coletar tempos médios em cada situação
    todos_tempos_medios = defaultdict(list)
    for r in todos_resultados.values():
        if r.get('status') == 'SUCESSO' and r.get('tempos_medios'):
            for situacao, tempo in r.get('tempos_medios', {}).items():
                todos_tempos_medios[situacao].append(tempo)
    
    # Calcular média dos tempos médios
    tempos_medios_consolidados = {
        situacao: sum(tempos) / len(tempos) 
        for situacao, tempos in todos_tempos_medios.items() 
        if len(tempos) > 0
    }
    
    # Identificar colaboradores com melhor e pior qualidade
    colaboradores_validos = [r for r in todos_resultados.values() 
                            if r.get('status') == 'SUCESSO' and 'score_qualidade' in r]
    
    if colaboradores_validos:
        melhor_colaborador = max(colaboradores_validos, key=lambda x: x.get('score_qualidade', 0))
        pior_colaborador = min(colaboradores_validos, key=lambda x: x.get('score_qualidade', 0))
    else:
        melhor_colaborador = None
        pior_colaborador = None
    
    # Coletar caminhos dos gráficos gerados
    graficos_gerados = [
        r.get('grafico_path') for r in todos_resultados.values()
        if r.get('status') == 'SUCESSO' and r.get('grafico_path')
    ]
    
    # Gerar relatório
    relatorio = []
    relatorio.append("=" * 80)
    relatorio.append("RELATÓRIO DE MELHORIAS NA QUALIDADE DA ANÁLISE")
    relatorio.append("=" * 80)
    relatorio.append(f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    relatorio.append(f"Total de Colaboradores Analisados: {total_colaboradores}")
    relatorio.append(f"Colaboradores com Problemas Identificados: {colaboradores_com_problemas} ({colaboradores_com_problemas/total_colaboradores*100:.1f}%)")
    relatorio.append(f"Gráficos Gerados: {len(graficos_gerados)}")
    relatorio.append("")
    
    relatorio.append("-" * 80)
    relatorio.append("PROBLEMAS MAIS COMUNS")
    relatorio.append("-" * 80)
    for problema, contagem in contagem_problemas.most_common(5):
        relatorio.append(f"• {problema}: {contagem} ocorrências ({contagem/total_colaboradores*100:.1f}% dos colaboradores)")
    relatorio.append("")
    
    relatorio.append("-" * 80)
    relatorio.append("SUGESTÕES DE MELHORIA")
    relatorio.append("-" * 80)
    for sugestao, contagem in contagem_sugestoes.most_common(5):
        relatorio.append(f"• {sugestao}")
    relatorio.append("")
    
    if todas_transicoes:
        relatorio.append("-" * 80)
        relatorio.append("ANÁLISE DE TRANSIÇÕES DE ESTADO")
        relatorio.append("-" * 80)
        relatorio.append("Transições mais comuns entre estados:")
        for transicao, contagem in sorted(todas_transicoes.items(), key=lambda x: x[1], reverse=True)[:5]:
            relatorio.append(f"• {transicao}: {contagem} ocorrências")
        relatorio.append("")
    
    if tempos_medios_consolidados:
        relatorio.append("-" * 80)
        relatorio.append("TEMPO MÉDIO EM CADA SITUAÇÃO")
        relatorio.append("-" * 80)
        for situacao, tempo in sorted(tempos_medios_consolidados.items(), key=lambda x: x[1], reverse=True):
            relatorio.append(f"• {situacao}: {tempo:.1f} dias")
        relatorio.append("")
    
    if melhor_colaborador and pior_colaborador:
        relatorio.append("-" * 80)
        relatorio.append("MELHORES PRÁTICAS")
        relatorio.append("-" * 80)
        relatorio.append(f"Colaborador com Melhor Qualidade: {melhor_colaborador['colaborador']}")
        relatorio.append(f"Score de Qualidade: {melhor_colaborador['score_qualidade']:.1f}/100")
        relatorio.append(f"Taxa de Preenchimento: {melhor_colaborador['taxa_preenchimento']:.1f}%")
        relatorio.append(f"Taxa de Padronização: {melhor_colaborador['taxa_padronizacao']:.1f}%")
        relatorio.append(f"Consistência Diária: {melhor_colaborador['consistencia_diaria']:.1f}%")
        
        if melhor_colaborador.get('analise_transicoes'):
            relatorio.append("Transições de Estado:")
            for transicao, contagem in sorted(melhor_colaborador['analise_transicoes'].items(), 
                                             key=lambda x: x[1], reverse=True)[:3]:
                relatorio.append(f"  - {transicao}: {contagem} ocorrências")
        
        if melhor_colaborador.get('grafico_path'):
            relatorio.append(f"Gráfico de Distribuição: {melhor_colaborador['grafico_path']}")
        
        relatorio.append("")
        
        relatorio.append("-" * 80)
        relatorio.append("OPORTUNIDADES DE MELHORIA")
        relatorio.append("-" * 80)
        relatorio.append(f"Colaborador com Maior Oportunidade: {pior_colaborador['colaborador']}")
        relatorio.append(f"Score de Qualidade: {pior_colaborador['score_qualidade']:.1f}/100")
        relatorio.append(f"Taxa de Preenchimento: {pior_colaborador['taxa_preenchimento']:.1f}%")
        relatorio.append(f"Taxa de Padronização: {pior_colaborador['taxa_padronizacao']:.1f}%")
        relatorio.append(f"Consistência Diária: {pior_colaborador['consistencia_diaria']:.1f}%")
        relatorio.append("Problemas Identificados:")
        for problema in pior_colaborador.get('problemas', []):
            relatorio.append(f"  - {problema}")
        relatorio.append("Sugestões:")
        for sugestao in pior_colaborador.get('sugestoes', []):
            relatorio.append(f"  - {sugestao}")
    
    relatorio.append("")
    relatorio.append("-" * 80)
    relatorio.append("RECOMENDAÇÕES GERAIS PARA MELHORAR A QUALIDADE DA ANÁLISE")
    relatorio.append("-" * 80)
    relatorio.append("1. Padronizar os valores da coluna SITUAÇÃO para facilitar análises comparativas")
    relatorio.append("2. Garantir que todos os registros tenham a situação preenchida")
    relatorio.append("3. Manter uma frequência consistente de atualizações diárias")
    relatorio.append("4. Adicionar timestamps para cada atualização de status")
    relatorio.append("5. Implementar validação de dados na entrada para evitar erros de digitação")
    relatorio.append("6. Criar um dicionário de termos padronizados para referência dos colaboradores")
    relatorio.append("7. Realizar treinamentos periódicos sobre a importância da qualidade dos dados")
    relatorio.append("8. Implementar alertas automáticos para registros com dados incompletos")
    
    if graficos_gerados:
        relatorio.append("")
        relatorio.append("-" * 80)
        relatorio.append("GRÁFICOS GERADOS")
        relatorio.append("-" * 80)
        relatorio.append("Os seguintes gráficos foram gerados para análise visual:")
        for grafico in graficos_gerados[:10]:  # Limitar a 10 gráficos para não sobrecarregar o relatório
            relatorio.append(f"• {grafico}")
        if len(graficos_gerados) > 10:
            relatorio.append(f"... e mais {len(graficos_gerados) - 10} gráficos")
    
    relatorio.append("")
    relatorio.append("=" * 80)
    
    return "\n".join(relatorio)

def main():
    """Função principal para executar a análise em paralelo"""
    print("Iniciando análise paralela dos arquivos Excel...")
    
    # Definir caminhos dos arquivos
    arquivo_julio = "(JULIO) LISTAS INDIVIDUAIS.xlsx"
    arquivo_leandro = "(LEANDRO_ADRIANO) LISTAS INDIVIDUAIS.xlsx"
    
    print(f"Analisando arquivo: {arquivo_julio}")
    resultados_julio = analisar_arquivo_paralelo(arquivo_julio)
    
    print(f"Analisando arquivo: {arquivo_leandro}")
    resultados_leandro = analisar_arquivo_paralelo(arquivo_leandro)
    
    # Gerar relatório de melhorias
    relatorio = gerar_relatorio_melhorias(resultados_julio, resultados_leandro)
    
    # Salvar relatório em arquivo
    nome_arquivo = f"relatorio_melhorias_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    with open(nome_arquivo, 'w', encoding='utf-8') as f:
        f.write(relatorio)
    
    print(f"Relatório de melhorias salvo em: {nome_arquivo}")
    print(relatorio)
    
    return {
        'resultados_julio': resultados_julio,
        'resultados_leandro': resultados_leandro,
        'relatorio': relatorio
    }

if __name__ == "__main__":
    main()
