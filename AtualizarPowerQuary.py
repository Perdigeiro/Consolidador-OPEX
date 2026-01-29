import win32com.client as win32
import ctypes
import os
import time

def mostrar_popup(titulo, mensagem, estilo):
    """
    Estilos:
    0 : OK
    1 : OK | Cancelar
    16: Ícone de Erro Crítico
    48: Ícone de Alerta
    64: Ícone de Informação
    """
    ctypes.windll.user32.MessageBoxW(0, mensagem, titulo, estilo)

def atualizar_excel(caminho_arquivo):
    # Caminho do arquivo (Note o 'r' antes das aspas para ler caracteres especiais)

    # Verifica se o arquivo existe antes de tentar abrir
    if not os.path.exists(caminho_arquivo):
        mostrar_popup("Erro", f"Arquivo não encontrado no caminho:\n{caminho_arquivo}", 16)
        return

    excel_app = None
    workbook = None

    try:
        # Inicia a instância do Excel
        excel_app = win32.Dispatch("Excel.Application")
        
        # Define como invisível (segundo plano) e desativa alertas (como "Deseja salvar?")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False

        # Abre a pasta de trabalho
        workbook = excel_app.Workbooks.Open(caminho_arquivo)

        # Comando para Atualizar Tudo (Refresh All) - Dispara o Power Query
        workbook.RefreshAll()

        # CRUCIAL: Espera as queries assíncronas terminarem antes de salvar/fechar
        # Isso garante que o Power Query termine de baixar os dados da rede
        excel_app.CalculateUntilAsyncQueriesDone()

        # Salva e fecha
        workbook.Save()
        workbook.Close()

        nome_planilha = caminho_arquivo.split("\\")[-1]
        
        # Popup de Sucesso
        mostrar_popup("Sucesso", f"A base de dados {nome_planilha} foi atualizada com sucesso!", 64)

    except Exception as e:
        # Popup de Erro com a descrição do problema
        mostrar_popup("Erro na Atualização", f"Ocorreu um erro ao atualizar:\n{str(e)}", 16)

    finally:
        # Garante que o Excel feche, mesmo se der erro, para não travar a memória
        if excel_app:
            try:
                excel_app.Quit()
            except:
                pass

if __name__ == "__main__":
    atualizar_excel(r"C:\Users\matheus.augusto\OneDrive - Grupo Fleury\Planejamento Financeiro - TI&Telecom _ - Documentos\RELATÓRIO\[ORÇ] CAPEX & OPEX 2026.xlsx")