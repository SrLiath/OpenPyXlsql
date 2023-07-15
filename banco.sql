-- Criação da tabela Produto_Cartao
CREATE TABLE Produto_Cartao (
    id_produtocartao INT PRIMARY KEY,
    nm_ProdutoCartao VARCHAR(100)
);

-- Criação da tabela Contratante
CREATE TABLE Contratante (
    id_Contratante INT PRIMARY KEY,
    no_CNPJ VARCHAR(20)
);

-- Criação da tabela Cartao
CREATE TABLE Cartao (
    id_Cartao INT PRIMARY KEY,
    no_SequencialEmissao VARCHAR(20),
    id_Contratante INT,
    id_PessoaConta INT,
    id_tiposituacaocartao INT
);



