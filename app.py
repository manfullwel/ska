from flask import Flask, render_template, jsonify, request
from auditoria_dados import AuditorDados
from analise_360 import Analise360
import os

app = Flask(__name__)

class DataAnalyticsSAAS:
    def __init__(self):
        self.auditor = AuditorDados()
        self.analise360 = Analise360()
        self.setup_default_files()
        
    def setup_default_files(self):
        """Configura os arquivos padrão para análise"""
        self.arquivos = {
            'JULIO': os.path.join(os.getcwd(), "(JULIO) LISTAS INDIVIDUAIS.xlsx"),
            'LEANDRO': os.path.join(os.getcwd(), "(LEANDRO_ADRIANO) LISTAS INDIVIDUAIS.xlsx")
        }
        self.auditor.arquivos = self.arquivos
        self.analise360.configurar_arquivos(self.arquivos)
        
    def get_section_data(self, section, filters=None):
        """Obtém dados baseados na seção e filtros"""
        if section == 'acordo':
            return self.get_acordo_data(filters)
        elif section == 'diario':
            return self.get_daily_report(filters)
        else:
            return self.get_general_report(filters)
            
    def get_acordo_data(self, filters):
        """Obtém dados relacionados a acordos"""
        try:
            df_ranking = self.analise360.gerar_ranking()
            acordos_data = {
                'kpis': {
                    'total': len(df_ranking),
                    'pending': df_ranking['Casos Pendentes'].sum(),
                    'resolved': df_ranking['Casos Processados'].sum(),
                    'rate': round(df_ranking['Taxa Eficiência'].mean(), 2)
                },
                'charts': {
                    'status': df_ranking['Status'].value_counts().to_dict(),
                    'timeline': df_ranking.groupby('Data')['Casos Processados'].sum().to_dict()
                },
                'table': df_ranking.to_dict('records')
            }
            return acordos_data
        except Exception as e:
            return {'error': str(e)}
            
    def get_daily_report(self, filters):
        """Obtém relatório diário"""
        try:
            dados_diarios = {}
            for nome, arquivo in self.arquivos.items():
                df = self.auditor.relatorio_completo.get(nome, {}).get('abas', {})
                if df:
                    for aba, dados in df.items():
                        dados_diarios[f"{nome}_{aba}"] = {
                            'total': dados['total_linhas'],
                            'nulos': dados['valores_nulos'],
                            'problemas': dados['problemas_detectados']
                        }
            return dados_diarios
        except Exception as e:
            return {'error': str(e)}
            
    def get_general_report(self, filters):
        """Obtém relatório geral"""
        try:
            return {
                'auditor': self.auditor.relatorio_completo,
                'analise360': self.analise360.gerar_ranking().to_dict('records')
            }
        except Exception as e:
            return {'error': str(e)}

# Instância global do SAAS
saas = DataAnalyticsSAAS()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/data', methods=['POST'])
def get_data():
    filters = request.json
    section = filters.get('section', 'geral')
    return jsonify(saas.get_section_data(section, filters))

@app.route('/api/update_title/<section>')
def update_title(section):
    titles = {
        'acordo': 'Gestão de Acordos',
        'diario': 'Relatório Diário',
        'geral': 'Relatório Geral'
    }
    return jsonify({'title': titles.get(section, 'DataAnalytics SAAS')})

if __name__ == '__main__':
    app.run(debug=True, port=5000)