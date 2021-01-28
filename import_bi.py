import pandas as pd
import psycopg2
##!! IMPORTANTE - - CORRIGJIR CAMPO NA TABELA Q ESTA COM  " ( ".

#con = psycopg2.connect(database='cap',host='10.71.16.108', user='postgres', password='GESTAOPROJETOS')
cur = conn.cursor()
limpar = 'truncate tb_bi'  #comando para Limpar a tabela
cur.execute(limpar)
print('Conexao Blezinha')

df = pd.read_excel('SeuArquivo.xlsx','NomePlanilha') # sintaxe pd.read_excel('caminho_do_arquivo completo",'NomePLanilha')
ordem_de_servico = df.iloc[:,0]
macro_area = df.iloc[:,1]
regional = df.iloc[:,2]
resp = df.iloc[: ,3]
municipio = df.iloc[: ,4]
tipo_pasta = df.iloc[: ,5]
empresa_construtora = df.iloc[: ,6]
controle = df.iloc[: ,7]
tipo_doc = df.iloc[: ,8]
numero = df.iloc[: ,9]
visita = df.iloc[: ,10]
baremo = df.iloc[: ,11]
data_envio = df.iloc[: ,12]
ano = df.iloc[: ,13]
mes = df.iloc[: ,14]
numero_da_semana_envio = df.iloc[: ,15]
empresa = df.iloc[: ,16]
data_retorno = df.iloc[: ,17]
mes_retorno = df.iloc[: ,18]
numero_da_semana_retorno = df.iloc[: ,19]
tempo_de_fiscalizacao = df.iloc[: ,20]
prazo = df.iloc[: ,21]
nao_conforme_defeito_grave = df.iloc[: ,22]
itens_fiscalizados = df.iloc[: ,23]
itens_reprovados = df.iloc[: ,24]
status = df.iloc[: ,25]
aging = df.iloc[: ,26]
latitude = df.iloc[: ,27]
longitude = df.iloc[: ,28]
estado = df.iloc[: ,29]
pais = df.iloc[: ,30]
nivel_de_defeito = df.iloc[: ,31]
tempo_minimo_para_fiscalizacao = df.iloc[: ,32]
tempo_maximo_para_fiscalizacao = df.iloc[: ,33]
contador_fiscalizado = df.iloc[: ,34]
contador_itens_reprovados = df.iloc[: ,35]
contador_nao_conformidade = df.iloc[: ,36]
titulo_descricao = df.iloc[: ,37]
prj_assunto = df.iloc[: ,38]
projeto_de_investimento = df.iloc[: ,39]
contrato_sap = df.iloc[: ,40]
nomenclatura_contrato = df.iloc[: ,41]
descricao_contrato = df.iloc[:, 42]

lista = df.values.tolist()
for i, row in enumerate(lista):
    sql = (f"INSERT INTO tb_bi(ordem_de_servico,macro_area,regional,resp,"
           f"municipio,tipo_pasta,empresa_construtora,controle,tipo_doc,numero,visita,baremo,"
           f"data_envio,ano,mes,numero_da_semana_envio,empresa, data_retorno,mes_retorno,"
           f"numero_da_semana_retorno ,tempo_de_fiscalizacao, prazo,"
           f"nao_conforme_defeito_grave,itens_fiscalizados,itens_reprovados,status,aging,"
           f"latitude,longitude,estado,pais,nivel_de_defeito,tempo_minimo_para_fiscalizacao,"
           f"tempo_maximo_para_fiscalizacao, contador_fiscalizado,contador_itens_reprovados,"
           f"contador_nao_conformidade,titulo_descricao,prj_assunto,"
           f"projeto_de_investimento,contrato_sap,nomenclatura_contrato,"
           f"descricao_contrato) VALUES ('{ordem_de_servico[i]}',"
           f"'{macro_area[i]}','{regional[i]}','{resp[i]}',"
           f"'{municipio[i]}','{tipo_pasta[i]}','{empresa_construtora[i]}','{controle[i]}',"
           f"'{tipo_doc[i]}',{numero[i]}, {visita[i]},'{baremo[i]}','{data_envio[i]}',"
           f"'{ano[i]}','{mes[i]}',{numero_da_semana_envio[i]},'{empresa[i]}',"
           f"'{data_retorno[i]}','{mes_retorno[i]}','{numero_da_semana_retorno[i]}',"
           f"{tempo_de_fiscalizacao[i]},'{prazo[i]}',{nao_conforme_defeito_grave[i]},"
           f"{itens_fiscalizados[i]},{itens_reprovados[i]},'{status[i]}',"
           f"'{aging[i]}','{latitude[i]}','{longitude[i]}','{estado[i]}',"
           f"'{pais[i]}','{nivel_de_defeito[i]}',{tempo_minimo_para_fiscalizacao[i]},"
           f"{tempo_maximo_para_fiscalizacao[i]},{contador_fiscalizado[i]},"
           f"{contador_itens_reprovados[i]},{contador_nao_conformidade[i]} ,"
           f"'{titulo_descricao[i]}','{prj_assunto[i]}','{projeto_de_investimento[i]}',"
           f"'{contrato_sap[i]}','{nomenclatura_contrato[i]}','{descricao_contrato[i]}')")
    cur.execute(sql)
conn.commit()
conn.close()
print("TB Bi gravado com sucesso")