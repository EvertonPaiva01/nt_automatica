from docx import Document

# Abre o documento .docx
doc = Document("NT XX.2024 - Vistoria Pólo de Carnaval 'Nome'.docx")

# # Define o nome do trabalhador
nome_polo = "POLO TESTE"
logradouro = "Av. Norte Miguel Arraes de Alencar"
bairro = "Casa Amarela"
data = "08"
hora = "08:00"
nome_tecnico = "Everton"
cargo = "Programador"

visto_info = {
    1:"A infraestrutura não possui ART.",

    2:"O projeto elétrico com os equipamentos a serem instalados não foram encaminhado para a EMLURB juntamente a ART do responsável técnico pelas instalações?",

    3:"O responsável técnico não estava presente durante a montagem",

    4:"O fornecimento da energia não foi realizado pela Neoenergia PE",

    5:"Existem cargas conectadas diretamente a rede elétrica de IP",

    6:"A instalação não Disjuntor Termomagnético",

    7:"A instalação não possui DPS",

    8:"A instalação não possui DR",

    9:"O DR instalado não é de alta sensibilidade (30mA)",

    10:"As partes metálicas que são componentes dos sistemas elétricos não estão aterradas",

    11:"As partes metálicas que são próximas as instalações elétricas estão aterradas?",

    12:"O aterramento está executado com vara Copperweld 5/8x2,40m, fincada no solo com conector do tipo GAR/GTDU ou solda exotérmica, utilizando cabo de cobre 0,6/1kV de cor verde com bitola mínima de 16mm²?",

    13:"O aterramento das massas foi respeitado?",

    14:"Existem cabeamentos elétricos expostos ao tempo, intempéries, passando pelo chão ou amarrados em estruturas metálicas?",

    15:"Caso o item 15 seja verdadeiro, foi utilizado um trilho passa cabo protegendo os cabeamentos?",

    16:"Existem projetores/refletores ou qualquer tipo de iluminação provisória, em postes metálicos(exclusivos para iluminação pública) ou árvores?",

    17:"Os equipamentos usados para a realização das instalações elétricas provisórias são de boa qualidade?",

    18:"Existem elementos (bandeira, barracas, refletores, fiações, entre outros) nos postes e demais equipamentos componentes do sistema de iluminação pública?",

    19:"As distâncias para a AT e MT estão respeitadas(conforme ABNT e Neoenergia)?",

    20:"As caixas de passagem subterrânea estão violadas?",

    21:"Os medidores da Neoenergia estão violados?",

    22:"Existem cordões luminosos ou gambiarras cruzando vias, grades ou elementos metálicos pertencentes às estruturas dos prédios, comércios, residências, etc?"
}

def table_search():
    # Busca na Tabela
    for table in doc.tables:
        # Itera sobre todas as linhas da tabela
        for row in table.rows:
            # Itera sobre todas as células da linha
            for cell in row.cells:
                # Itera sobre todos os parágrafos dentro da célula
                for paragraph in cell.paragraphs:
                    # Verifica se o trecho desejado está no texto do parágrafo
                    if "‘Nome do Pólo’" in paragraph.text:
                        # Substitui "Nome do técnico." pelo nome do trabalhador
                        print(paragraph.text)
                        paragraph.text = paragraph.text.replace("‘Nome do Pólo’", nome_polo)



# Salva as alterações no documento
doc.save("seu_documento_modificado.docx")