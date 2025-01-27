CREATE TABLE Dissertacoes (
    id INT IDENTITY(1,1) PRIMARY KEY, -- ID auto-incrementado
    titulo NVARCHAR(MAX) NOT NULL, -- Título do trabalho
    autor NVARCHAR(MAX) NOT NULL, -- Nome do autor
    ano INT NOT NULL, -- Ano de publicação
    arquivo VARBINARY(MAX) NOT NULL, -- Arquivo PDF em formato binário
    tipo_arquivo NVARCHAR(50) NOT NULL, -- Tipo do arquivo (ex.: application/pdf)
    extensao NVARCHAR(10) NOT NULL, -- Extensão do arquivo (ex.: .pdf)
    nome_arquivo NVARCHAR(MAX) NOT NULL, -- Nome original do arquivo
    tamanho_arquivo BIGINT NOT NULL -- Tamanho do arquivo em bytes
);

pip install selenium
pip install pyautogui
pip install pyodbc

