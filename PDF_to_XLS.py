import streamlit as st
import tabula
import pandas as pd
from io import BytesIO
import tempfile
import os
import pdfplumber

@st.cache_data
def process_pdf(file):
    try:
        # Salvar o arquivo temporariamente para processamento
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(file.getvalue())
            tmp_path = tmp_file.name
        
        # Extrair tabelas usando pdfplumber
        tables = []
        with pdfplumber.open(tmp_path) as pdf:
            for page in pdf.pages:
                extracted_tables = page.extract_tables()
                if extracted_tables:
                    for table in extracted_tables:
                        # Converter para DataFrame
                        df = pd.DataFrame(table[1:], columns=table[0])
                        tables.append(df)
        
        # Limpar arquivo temporário
        os.unlink(tmp_path)
        
        return tables
    except Exception as e:
        raise Exception(f"Erro ao processar PDF: {str(e)}")

def main():
    st.title('Conversor de PDF para Excel')
    
    # Adicionar instruções
    st.write("Este aplicativo converte tabelas de arquivos PDF para Excel.")
    
    uploaded_file = st.file_uploader("Carregue seu arquivo PDF", type="pdf")

    if uploaded_file is not None:
        try:
            with st.spinner("Processando arquivo..."):
                tables = process_pdf(uploaded_file)
                
                if not tables:
                    st.warning("Nenhuma tabela encontrada no PDF.")
                else:
                    st.success(f"{len(tables)} tabelas encontradas!")
                    
                    # Criar arquivo Excel em memória
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for i, table in enumerate(tables, 1):
                            # Limpar e preparar dados antes de salvar
                            table = table.fillna('')  # Preencher valores nulos
                            table.to_excel(writer, sheet_name=f'Tabela_{i}', index=False)
                    
                    # Resetar o ponteiro do BytesIO
                    output.seek(0)
                    
                    # Configurar o botão de download
                    st.download_button(
                        label="Baixar arquivo Excel",
                        data=output.getvalue(),
                        file_name='tabelas_extraidas.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                    
                    # Mostrar pré-visualização com limitação de linhas
                    st.subheader("Pré-visualização das tabelas")
                    for i, table in enumerate(tables, 1):
                        with st.expander(f"Tabela {i}"):
                            st.dataframe(table.head(100))  # Mostra apenas as primeiras 100 linhas
                    
        except Exception as e:
            st.error(f"Ocorreu um erro: {str(e)}")

if __name__ == "__main__":
    main()