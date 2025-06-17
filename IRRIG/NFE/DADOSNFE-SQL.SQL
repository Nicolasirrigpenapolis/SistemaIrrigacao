/* Esquema para a criação do banco de dados da aplicação      */
/* Banco/Esquema: DADOSNFE                                    */


/*------------------------------------------------------------*/
/*        Exclusão de triggers/foreign key constraints        */
/*------------------------------------------------------------*/

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Baixa_contas_pagar_Manutenção_contas_pagar') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Baixa contas pagar]
        DROP CONSTRAINT [FK_Baixa_contas_pagar_Manutenção_contas_pagar]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Baixa_contas_receber_Manutenção_contas_receber') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Baixa contas receber]
        DROP CONSTRAINT [FK_Baixa_contas_receber_Manutenção_contas_receber]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Baixa_contas_receber_Tipo_da_Cobrança') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Baixa contas receber]
        DROP CONSTRAINT [FK_Baixa_contas_receber_Tipo_da_Cobrança]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Cheques_Pre_Bancos') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Cheques Pre]
        DROP CONSTRAINT [FK_Cheques_Pre_Bancos]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Cheques_Pre_Geral') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Cheques Pre]
        DROP CONSTRAINT [FK_Cheques_Pre_Geral]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Entrada_Contas_Pagar_Geral') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Entrada Contas Pagar]
        DROP CONSTRAINT [FK_Entrada_Contas_Pagar_Geral]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Entrada_Contas_Receber_Geral') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Entrada Contas Receber]
        DROP CONSTRAINT [FK_Entrada_Contas_Receber_Geral]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Geral_Municipios') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].Geral
        DROP CONSTRAINT [FK_Geral_Municipios]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Itens_da_OS_Produtos') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Itens da OS]
        DROP CONSTRAINT [FK_Itens_da_OS_Produtos]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Manutenção_contas_pagar_Geral') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Manutenção contas pagar]
        DROP CONSTRAINT [FK_Manutenção_contas_pagar_Geral]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Manutenção_contas_receber_Geral') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Manutenção contas receber]
        DROP CONSTRAINT [FK_Manutenção_contas_receber_Geral]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Modelo_do_Virabrequim_Marcas') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Modelo do Virabrequim]
        DROP CONSTRAINT [FK_Modelo_do_Virabrequim_Marcas]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Movimento_do_estoque_Geral') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Movimento do estoque]
        DROP CONSTRAINT [FK_Movimento_do_estoque_Geral]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Nota_fiscal_Geral') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Nota fiscal]
        DROP CONSTRAINT [FK_Nota_fiscal_Geral]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Nota_fiscal_Natureza_de_operação') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Nota fiscal]
        DROP CONSTRAINT [FK_Nota_fiscal_Natureza_de_operação]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Nota_fiscal_Classificação_fiscal') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Nota fiscal]
        DROP CONSTRAINT [FK_Nota_fiscal_Classificação_fiscal]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Ordem_de_Serviço_Geral') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Ordem de Serviço]
        DROP CONSTRAINT [FK_Ordem_de_Serviço_Geral]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Outras_saidas_estoque_Geral') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Outras saidas estoque]
        DROP CONSTRAINT [FK_Outras_saidas_estoque_Geral]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Produtos_Classificação_fiscal') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].Produtos
        DROP CONSTRAINT [FK_Produtos_Classificação_fiscal]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Produtos_Unidades') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].Produtos
        DROP CONSTRAINT [FK_Produtos_Unidades]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Produtos_da_nota_fiscal_Produtos') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Produtos da nota fiscal]
        DROP CONSTRAINT [FK_Produtos_da_nota_fiscal_Produtos]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Produtos_da_Venda_Produtos') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Produtos da Venda]
        DROP CONSTRAINT [FK_Produtos_da_Venda_Produtos]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Produtos_devolução_compra_Produtos') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Produtos devolução compra]
        DROP CONSTRAINT [FK_Produtos_devolução_compra_Produtos]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Produtos_movimento_estoque_Produtos') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Produtos movimento estoque]
        DROP CONSTRAINT [FK_Produtos_movimento_estoque_Produtos]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Serviços_da_Nota_Fiscal_Serviços') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Serviços da Nota Fiscal]
        DROP CONSTRAINT [FK_Serviços_da_Nota_Fiscal_Serviços]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Serviços_da_OS_Serviços') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Serviços da OS]
        DROP CONSTRAINT [FK_Serviços_da_OS_Serviços]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Serviços_da_Venda_Serviços') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Serviços da Venda]
        DROP CONSTRAINT [FK_Serviços_da_Venda_Serviços]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Vendas_Geral') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].Vendas
        DROP CONSTRAINT [FK_Vendas_Geral]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Virabrequim_Geral') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].Virabrequim
        DROP CONSTRAINT [FK_Virabrequim_Geral]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Virabrequim_Marcas') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].Virabrequim
        DROP CONSTRAINT [FK_Virabrequim_Marcas]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_Virabrequim_da_OS_Virabrequim') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[Virabrequim da OS]
        DROP CONSTRAINT [FK_Virabrequim_da_OS_Virabrequim]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_PW~Usuarios_PW~Grupos') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[PW~Usuarios]
        DROP CONSTRAINT [FK_PW~Usuarios_PW~Grupos]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('FK_PW~Tabelas_PW~Grupos') AND sysstat & 0xf = 11)
  ALTER TABLE [dbo].[PW~Tabelas]
        DROP CONSTRAINT [FK_PW~Tabelas_PW~Grupos]
 GO


/*------------------------------------------------------------*/
/*                     Exclusão de tabelas                    */
/*------------------------------------------------------------*/

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Baixa Casco]') AND sysstat & 0xf = 3)
  DROP TABLE [Baixa Casco]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Baixa contas pagar]') AND sysstat & 0xf = 3)
  DROP TABLE [Baixa contas pagar]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Baixa contas receber]') AND sysstat & 0xf = 3)
  DROP TABLE [Baixa contas receber]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('Bancos') AND sysstat & 0xf = 3)
  DROP TABLE Bancos
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Carta de correção nfe]') AND sysstat & 0xf = 3)
  DROP TABLE [Carta de correção nfe]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Centro de Custo]') AND sysstat & 0xf = 3)
  DROP TABLE [Centro de Custo]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Cheques Pre]') AND sysstat & 0xf = 3)
  DROP TABLE [Cheques Pre]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Classificação fiscal]') AND sysstat & 0xf = 3)
  DROP TABLE [Classificação fiscal]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Comissão Auxiliar]') AND sysstat & 0xf = 3)
  DROP TABLE [Comissão Auxiliar]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Devolução de compra]') AND sysstat & 0xf = 3)
  DROP TABLE [Devolução de compra]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('Empresa') AND sysstat & 0xf = 3)
  DROP TABLE Empresa
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Entrada Contas Pagar]') AND sysstat & 0xf = 3)
  DROP TABLE [Entrada Contas Pagar]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Entrada Contas Receber]') AND sysstat & 0xf = 3)
  DROP TABLE [Entrada Contas Receber]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Finalidade da nfe]') AND sysstat & 0xf = 3)
  DROP TABLE [Finalidade da nfe]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('Geral') AND sysstat & 0xf = 3)
  DROP TABLE Geral
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('Icms') AND sysstat & 0xf = 3)
  DROP TABLE Icms
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Itens da OS]') AND sysstat & 0xf = 3)
  DROP TABLE [Itens da OS]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Manutenção Cascos]') AND sysstat & 0xf = 3)
  DROP TABLE [Manutenção Cascos]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Manutenção contas pagar]') AND sysstat & 0xf = 3)
  DROP TABLE [Manutenção contas pagar]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Manutenção contas receber]') AND sysstat & 0xf = 3)
  DROP TABLE [Manutenção contas receber]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('Marcas') AND sysstat & 0xf = 3)
  DROP TABLE Marcas
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Modelo do Virabrequim]') AND sysstat & 0xf = 3)
  DROP TABLE [Modelo do Virabrequim]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Movimento do estoque]') AND sysstat & 0xf = 3)
  DROP TABLE [Movimento do estoque]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('Municipios') AND sysstat & 0xf = 3)
  DROP TABLE Municipios
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Natureza de operação]') AND sysstat & 0xf = 3)
  DROP TABLE [Natureza de operação]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[NFe devolução]') AND sysstat & 0xf = 3)
  DROP TABLE [NFe devolução]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Nota fiscal]') AND sysstat & 0xf = 3)
  DROP TABLE [Nota fiscal]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Ordem de Serviço]') AND sysstat & 0xf = 3)
  DROP TABLE [Ordem de Serviço]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Outras saidas estoque]') AND sysstat & 0xf = 3)
  DROP TABLE [Outras saidas estoque]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Parametros da nfe]') AND sysstat & 0xf = 3)
  DROP TABLE [Parametros da nfe]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Parametros da OS]') AND sysstat & 0xf = 3)
  DROP TABLE [Parametros da OS]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Parcelas da os]') AND sysstat & 0xf = 3)
  DROP TABLE [Parcelas da os]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Parcelas da Venda]') AND sysstat & 0xf = 3)
  DROP TABLE [Parcelas da Venda]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Parcelas dev compra]') AND sysstat & 0xf = 3)
  DROP TABLE [Parcelas dev compra]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Parcelas Entradas Pagar]') AND sysstat & 0xf = 3)
  DROP TABLE [Parcelas Entradas Pagar]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Parcelas Entradas Receber]') AND sysstat & 0xf = 3)
  DROP TABLE [Parcelas Entradas Receber]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Parcelas movimento estoque]') AND sysstat & 0xf = 3)
  DROP TABLE [Parcelas movimento estoque]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Parcelas nota fiscal]') AND sysstat & 0xf = 3)
  DROP TABLE [Parcelas nota fiscal]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Parcelas outras saidas]') AND sysstat & 0xf = 3)
  DROP TABLE [Parcelas outras saidas]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('Produtos') AND sysstat & 0xf = 3)
  DROP TABLE Produtos
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Produtos da nota fiscal]') AND sysstat & 0xf = 3)
  DROP TABLE [Produtos da nota fiscal]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Produtos da Venda]') AND sysstat & 0xf = 3)
  DROP TABLE [Produtos da Venda]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Produtos devolução compra]') AND sysstat & 0xf = 3)
  DROP TABLE [Produtos devolução compra]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Produtos movimento estoque]') AND sysstat & 0xf = 3)
  DROP TABLE [Produtos movimento estoque]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Produtos saidas estoque]') AND sysstat & 0xf = 3)
  DROP TABLE [Produtos saidas estoque]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('Serviços') AND sysstat & 0xf = 3)
  DROP TABLE Serviços
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Serviços da Nota Fiscal]') AND sysstat & 0xf = 3)
  DROP TABLE [Serviços da Nota Fiscal]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Serviços da OS]') AND sysstat & 0xf = 3)
  DROP TABLE [Serviços da OS]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Serviços da Venda]') AND sysstat & 0xf = 3)
  DROP TABLE [Serviços da Venda]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Tipo da Cobrança]') AND sysstat & 0xf = 3)
  DROP TABLE [Tipo da Cobrança]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('Unidades') AND sysstat & 0xf = 3)
  DROP TABLE Unidades
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('Vendas') AND sysstat & 0xf = 3)
  DROP TABLE Vendas
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('Virabrequim') AND sysstat & 0xf = 3)
  DROP TABLE Virabrequim
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[Virabrequim da OS]') AND sysstat & 0xf = 3)
  DROP TABLE [Virabrequim da OS]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[PW~Grupos]') AND sysstat & 0xf = 3)
  DROP TABLE [PW~Grupos]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[PW~Tabelas]') AND sysstat & 0xf = 3)
  DROP TABLE [PW~Tabelas]
 GO

 IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('[PW~Usuarios]') AND sysstat & 0xf = 3)
  DROP TABLE [PW~Usuarios]
 GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                         Baixa Casco                        */
/*------------------------------------------------------------*/

 CREATE TABLE [Baixa Casco] (
       [Sequencia da baixa]             int                 DEFAULT 0 NOT NULL, 
       [Data da baixa]                  datetime            NULL, 
       Baixado                          bit                 DEFAULT 0 NOT NULL, 
       [Sequencia da Manutenção]        int                 DEFAULT 0 NOT NULL, 
       Teste                            varchar (3)         DEFAULT '' NOT NULL, 
          CONSTRAINT Seq_Casco PRIMARY KEY CLUSTERED
          (
                  [Sequencia da baixa]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                     Baixa contas pagar                     */
/*------------------------------------------------------------*/

 CREATE TABLE [Baixa contas pagar] (
       [Sequencia da baixa]             int                 NOT NULL     IDENTITY (1,  1 ) NOT FOR REPLICATION, 
       [Sequencia da Manutenção]        int                 DEFAULT 0 NOT NULL, 
       [Data da baixa]                  datetime            NULL, 
       [Valor pago]                     decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Valor do juros]                 decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Valor do desconto]              decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Numero do cheque]               varchar (10)        DEFAULT '' NOT NULL, 
       Historico                        text                DEFAULT '' NOT NULL, 
       [Sequencia da Cobrança]          smallint            DEFAULT 0 NOT NULL, 
       Teste                            varchar (3)         DEFAULT '' NOT NULL, 
          CONSTRAINT [Sequencia da baixa pg] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da baixa]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                    Baixa contas receber                    */
/*------------------------------------------------------------*/

 CREATE TABLE [Baixa contas receber] (
       [Sequencia da baixa]             int                 NOT NULL     IDENTITY (1,  1 ) NOT FOR REPLICATION, 
       [Sequencia da Manutenção]        int                 DEFAULT 0 NOT NULL, 
       [Data da baixa]                  datetime            NULL, 
       [Valor pago]                     decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Valor do juros]                 decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Valor do desconto]              decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Numero do cheque]               varchar (10)        DEFAULT '' NOT NULL, 
       Historico                        text                DEFAULT '' NOT NULL, 
       Bloqueado                        bit                 DEFAULT 0 NOT NULL, 
       [Sequencia da Cobrança]          smallint            DEFAULT 0 NOT NULL, 
       [Data da Alteração]              datetime            NULL, 
       [Hora da Alteração]              datetime            NULL, 
       [Usuario da Alteração]           varchar (30)        DEFAULT '' NOT NULL, 
       Teste                            varchar (3)         DEFAULT '' NOT NULL, 
          CONSTRAINT Seq_Bx_Receber PRIMARY KEY CLUSTERED
          (
                  [Sequencia da baixa]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                           Bancos                           */
/*------------------------------------------------------------*/

 CREATE TABLE Bancos (
       [Codigo do Banco]                smallint            DEFAULT 0 NOT NULL, 
       [Nome do Banco]                  varchar (40)        DEFAULT '' NOT NULL, 
          CONSTRAINT [Codigo do Banco] PRIMARY KEY CLUSTERED
          (
                  [Codigo do Banco]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                    Carta de correção nfe                   */
/*------------------------------------------------------------*/

 CREATE TABLE [Carta de correção nfe] (
       [Sequencia da correção]          int                 DEFAULT 0 NOT NULL, 
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       [Numero da correção]             smallint            DEFAULT 0 NOT NULL, 
       [Justificativa cce]              text                DEFAULT '' NOT NULL, 
       [Data correção]                  datetime            NULL, 
          CONSTRAINT [Sequencia da correção] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da correção], [Sequencia da nota fiscal]
          ) WITH FILLFACTOR = 90

 )
GO

 CREATE UNIQUE INDEX SeqCorrecao ON [Carta de correção nfe]
       ([Sequencia da nota fiscal], [Numero da correção]) WITH FILLFACTOR = 90
GO




/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                       Centro de Custo                      */
/*------------------------------------------------------------*/

 CREATE TABLE [Centro de Custo] (
       [Codigo do Centro]               smallint            DEFAULT 0 NOT NULL, 
       [Descrição do Centro]            varchar (40)        DEFAULT '' NOT NULL, 
       Teste                            varchar (3)         DEFAULT '' NOT NULL, 
          CONSTRAINT [Codigo do Centro] PRIMARY KEY CLUSTERED
          (
                  [Codigo do Centro]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                         Cheques Pre                        */
/*------------------------------------------------------------*/

 CREATE TABLE [Cheques Pre] (
       [Sequencia do Cheque]            int                 NOT NULL     IDENTITY (1,  1 ) NOT FOR REPLICATION, 
       Documento                        varchar (25)        DEFAULT '' NOT NULL, 
       [Codigo do Banco]                smallint            DEFAULT 0 NOT NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       Agencia                          varchar (10)        DEFAULT '' NOT NULL, 
       [Data da entrada]                datetime            NULL, 
       [Data Boa]                       datetime            NULL, 
       [Valor do Cheque]                decimal (10, 2)     DEFAULT 0 NOT NULL, 
       Compensado                       bit                 DEFAULT 0 NOT NULL, 
       [Data da Compessação]            datetime            NULL, 
       Historico                        text                DEFAULT '' NOT NULL, 
       [Vendedor Responsavel]           int                 DEFAULT 0 NOT NULL, 
       [Numero do cheque]               varchar (10)        DEFAULT '' NOT NULL, 
       [Emitente do Cheque]             varchar (60)        DEFAULT '' NOT NULL, 
       Devolvido                        bit                 DEFAULT 0 NOT NULL, 
       [Motivo da Devolução]            varchar (100)       DEFAULT '' NOT NULL, 
          CONSTRAINT [Sequencia do Cheque] PRIMARY KEY CLUSTERED
          (
                  [Sequencia do Cheque]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                    Classificação fiscal                    */
/*------------------------------------------------------------*/

 CREATE TABLE [Classificação fiscal] (
       [Sequencia da classificação]     int                 DEFAULT 0 NOT NULL, 
       Ncm                              varchar (8)         DEFAULT '' NOT NULL, 
       [Descrição do ncm]               varchar (80)        DEFAULT '' NOT NULL, 
       [Percentual de ipi]              decimal (8, 4)      DEFAULT 0 NOT NULL, 
       [Produto diferido]               bit                 DEFAULT 0 NOT NULL, 
       Iva                              decimal (7, 4)      DEFAULT 0 NOT NULL, 
       Cest                             varchar (7)         DEFAULT '' NOT NULL, 
       Inativo                          bit                 DEFAULT 0 NOT NULL, 
       [Redução da Bc]                  bit                 DEFAULT 0 NOT NULL, 
          CONSTRAINT [Sequencia da classificação] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da classificação]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                      Comissão Auxiliar                     */
/*------------------------------------------------------------*/

 CREATE TABLE [Comissão Auxiliar] (
       [Sequencia da Comissão]          int                 DEFAULT 0 NOT NULL, 
       Documento                        varchar (25)        DEFAULT '' NOT NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       [Data de entrada]                datetime            NULL, 
       [Data de vencimento]             datetime            NULL, 
       Parcela                          smallint            DEFAULT 0 NOT NULL, 
       [Valor da parcela]               decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Valor pago]                     decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Data da baixa]                  datetime            NULL, 
       [Valor restante]                 decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Sequencia do Vendedor]          int                 DEFAULT 0 NOT NULL, 
          CONSTRAINT [Sequencia da Comissão] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Comissão]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                     Devolução de compra                    */
/*------------------------------------------------------------*/

 CREATE TABLE [Devolução de compra] (
       [Sequencia da devolução]         int                 DEFAULT 0 NOT NULL, 
       [Sequencia do movimento]         int                 DEFAULT 0 NOT NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       [Data da entrada]                datetime            NULL, 
       [Data da devolução]              datetime            NULL, 
       Historico                        text                DEFAULT '' NOT NULL, 
       Cancelado                        bit                 DEFAULT 0 NOT NULL, 
       [Valor total dos produtos]       decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total do ipi]             decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total da devolução]       decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       [Data da Alteração]              datetime            NULL, 
       [Hora da Alteração]              datetime            NULL, 
       [Usuario da Alteração]           varchar (30)        DEFAULT '' NOT NULL, 
       Fechado                          bit                 DEFAULT 0 NOT NULL, 
       [Valor total icms st]            decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Forma de pagamento]             varchar (5)         DEFAULT '' NOT NULL, 
       [Valor do frete]                 decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor do desconto]              decimal (11, 2)     DEFAULT 0 NOT NULL
 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                           Empresa                          */
/*------------------------------------------------------------*/

 CREATE TABLE Empresa (
       [Razão social]                   varchar (60)        DEFAULT '' NOT NULL, 
       [Nome fantasia]                  varchar (60)        DEFAULT '' NOT NULL, 
       [Cpf do emitente]                varchar (14)        DEFAULT '' NOT NULL, 
       [Inscrição estadual]             varchar (20)        DEFAULT '' NOT NULL, 
       Cnpj                             varchar (18)        DEFAULT '' NOT NULL, 
       Rua                              varchar (60)        DEFAULT '' NOT NULL, 
       Nro                              varchar (8)         DEFAULT '' NOT NULL, 
       Bairro                           varchar (60)        DEFAULT '' NOT NULL, 
       Município                        varchar (60)        DEFAULT '' NOT NULL, 
       [Codigo do municipio]            int                 DEFAULT 0 NOT NULL, 
       Uf                               varchar (3)         DEFAULT '' NOT NULL, 
       Cep                              varchar (9)         DEFAULT '' NOT NULL, 
       Fone                             varchar (13)        DEFAULT '' NOT NULL, 
       [Aliquota do pis]                decimal (8, 4)      DEFAULT 0 NOT NULL, 
       [Aliquota do cofins]             decimal (8, 4)      DEFAULT 0 NOT NULL, 
       [Chave flexdocs]                 varchar (130)       DEFAULT '' NOT NULL, 
       Email                            varchar (255)       DEFAULT '' NOT NULL, 
       [Data Inicial da Ativação]       datetime            NULL, 
       [Senha Admin]                    int                 DEFAULT 0 NOT NULL, 
       [Aliquota do Iss]                decimal (5, 2)      DEFAULT 0 NOT NULL, 
       [Baixa Pagar Auto]               bit                 DEFAULT 0 NOT NULL, 
       [Baixa Receber Auto]             bit                 DEFAULT 0 NOT NULL, 
       [Incrição Municipal]             varchar (10)        DEFAULT '' NOT NULL, 
       [Emite NFe]                      bit                 DEFAULT 0 NOT NULL, 
       [Percentual da redução]          decimal (5, 2)      DEFAULT 0 NOT NULL, 
       [Não Controlar Estoque]          bit                 DEFAULT 0 NOT NULL, 
       [Impressão do Relatorio]         smallint            DEFAULT 0 NOT NULL, 
       LogoTipo                         image               NULL, 
       [Numeração Auto]                 bit                 DEFAULT 0 NOT NULL
 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                    Entrada Contas Pagar                    */
/*------------------------------------------------------------*/

 CREATE TABLE [Entrada Contas Pagar] (
       [Sequencia da entrada]           int                 NOT NULL     IDENTITY (1,  1 ) NOT FOR REPLICATION, 
       [Data de entrada]                datetime            NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       Documento                        varchar (25)        DEFAULT '' NOT NULL, 
       Historico                        text                DEFAULT '' NOT NULL, 
       [Forma de pagamento]             varchar (5)         DEFAULT '' NOT NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Data da Alteração]              datetime            NULL, 
       [Hora da Alteração]              datetime            NULL, 
       [Usuario da Alteração]           varchar (30)        DEFAULT '' NOT NULL, 
       [Tipo de conta]                  varchar (40)        DEFAULT '' NOT NULL, 
       Fechado                          bit                 DEFAULT 0 NOT NULL, 
       Cancelado                        bit                 DEFAULT 0 NOT NULL, 
       Teste                            varchar (3)         DEFAULT '' NOT NULL, 
          CONSTRAINT [Sequencia da entrada pg] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da entrada]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                   Entrada Contas Receber                   */
/*------------------------------------------------------------*/

 CREATE TABLE [Entrada Contas Receber] (
       [Sequencia da entrada]           int                 DEFAULT 0 NOT NULL, 
       [Data de entrada]                datetime            NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       Documento                        varchar (25)        DEFAULT '' NOT NULL, 
       Historico                        text                DEFAULT '' NOT NULL, 
       [Forma de pagamento]             varchar (5)         DEFAULT '' NOT NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Data da Alteração]              datetime            NULL, 
       [Hora da Alteração]              datetime            NULL, 
       [Usuario da Alteração]           varchar (30)        DEFAULT '' NOT NULL, 
       [Tipo de conta]                  varchar (40)        DEFAULT '' NOT NULL, 
       Fechado                          bit                 DEFAULT 0 NOT NULL, 
       Cancelado                        bit                 DEFAULT 0 NOT NULL, 
          CONSTRAINT [Sequencia da entrada] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da entrada]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                      Finalidade da nfe                     */
/*------------------------------------------------------------*/

 CREATE TABLE [Finalidade da nfe] (
       Codigo                           smallint            DEFAULT 0 NOT NULL, 
       Descrição                        varchar (80)        DEFAULT '' NOT NULL, 
          CONSTRAINT Codigo PRIMARY KEY CLUSTERED
          (
                  Codigo
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                            Geral                           */
/*------------------------------------------------------------*/

 CREATE TABLE Geral (
       Cliente                          bit                 DEFAULT 0 NOT NULL, 
       Transportadora                   bit                 DEFAULT 0 NOT NULL, 
       [Consumidor final]               bit                 DEFAULT 0 NOT NULL, 
       Isento                           bit                 DEFAULT 0 NOT NULL, 
       Inativo                          bit                 DEFAULT 0 NOT NULL, 
       [Sequencia do geral]             int                 NOT NULL     IDENTITY (1,  1 ) NOT FOR REPLICATION, 
       [Razão social]                   varchar (60)        DEFAULT '' NOT NULL, 
       [Nome fantasia]                  varchar (60)        DEFAULT '' NOT NULL, 
       Tipo                             smallint            DEFAULT 0 NOT NULL, 
       Endereço                         varchar (100)       DEFAULT '' NOT NULL, 
       Complemento                      varchar (30)        DEFAULT '' NOT NULL, 
       Nro                              varchar (8)         DEFAULT '' NOT NULL, 
       Bairro                           varchar (60)        DEFAULT '' NOT NULL, 
       [Sequencia do municipio]         int                 DEFAULT 0 NOT NULL, 
       Cep                              varchar (9)         DEFAULT '' NOT NULL, 
       Fone                             varchar (13)        DEFAULT '' NOT NULL, 
       Celular                          varchar (14)        DEFAULT '' NOT NULL, 
       [Codigo do suframa]              varchar (9)         DEFAULT '' NOT NULL, 
       [Cpf e cnpj]                     varchar (20)        DEFAULT '' NOT NULL, 
       [Rg e ie]                        varchar (20)        DEFAULT '' NOT NULL, 
       Observação                       text                DEFAULT '' NOT NULL, 
       [Data do cadastro]               datetime            NULL, 
       Fornecedor                       bit                 DEFAULT 0 NOT NULL, 
       Email                            varchar (255)       DEFAULT '' NOT NULL, 
       [Codigo da antt]                 varchar (20)        DEFAULT '' NOT NULL, 
       [Data de Nascimento]             datetime            NULL, 
       Despesa                          bit                 DEFAULT 0 NOT NULL, 
       Vendedor                         bit                 DEFAULT 0 NOT NULL, 
       Funcionario                      bit                 DEFAULT 0 NOT NULL, 
       Imposto                          bit                 DEFAULT 0 NOT NULL, 
       Produtor                         bit                 DEFAULT 0 NOT NULL, 
          CONSTRAINT [Sequencia do geral] PRIMARY KEY CLUSTERED
          (
                  [Sequencia do geral]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                            Icms                            */
/*------------------------------------------------------------*/

 CREATE TABLE Icms (
       [Sequencia do icms]              smallint            DEFAULT 0 NOT NULL, 
       Uf                               varchar (3)         DEFAULT '' NOT NULL, 
       Região                           varchar (20)        DEFAULT '' NOT NULL, 
       [Aliquota do icms]               decimal (5, 2)      DEFAULT 0 NOT NULL, 
       [Aliquota interna]               decimal (5, 2)      DEFAULT 0 NOT NULL, 
       [Codigo da uf]                   smallint            DEFAULT 0 NOT NULL, 
       Inativo                          bit                 DEFAULT 0 NOT NULL, 
          CONSTRAINT [Sequencia do icms] PRIMARY KEY CLUSTERED
          (
                  [Sequencia do icms]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                         Itens da OS                        */
/*------------------------------------------------------------*/

 CREATE TABLE [Itens da OS] (
       [Sequencia da Ordem]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia do produto]           bigint              DEFAULT 0 NOT NULL, 
       Quantidade                       decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor unitario]                 decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
          CONSTRAINT Seq_e_ItemOS PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Ordem], [Sequencia do produto]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                      Manutenção Cascos                     */
/*------------------------------------------------------------*/

 CREATE TABLE [Manutenção Cascos] (
       [Sequencia da Manutenção]        int                 NOT NULL     IDENTITY (1,  1 ) NOT FOR REPLICATION, 
       Documento                        varchar (25)        DEFAULT '' NOT NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia do Virabrequim]       int                 DEFAULT 0 NOT NULL, 
       [Sequencia da baixa]             int                 DEFAULT 0 NOT NULL, 
       [Data da baixa]                  datetime            NULL, 
       [Data de entrada]                datetime            NULL, 
       [Sequencia da Ordem]             int                 DEFAULT 0 NOT NULL, 
       Baixado                          bit                 DEFAULT 0 NOT NULL, 
       Teste                            varchar (3)         DEFAULT '' NOT NULL, 
          CONSTRAINT Seq_Manu_Casco PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Manutenção]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                   Manutenção contas pagar                  */
/*------------------------------------------------------------*/

 CREATE TABLE [Manutenção contas pagar] (
       [Sequencia da Manutenção]        int                 NOT NULL     IDENTITY (1,  1 ) NOT FOR REPLICATION, 
       Parcela                          smallint            DEFAULT 0 NOT NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       [Data de entrada]                datetime            NULL, 
       Historico                        text                DEFAULT '' NOT NULL, 
       [Forma de pagamento]             varchar (5)         DEFAULT '' NOT NULL, 
       [Data de vencimento]             datetime            NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Valor da parcela]               decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Valor pago]                     decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Valor do juros]                 decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Valor do desconto]              decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor restante]                 decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Tipo de conta]                  varchar (40)        DEFAULT '' NOT NULL, 
       [Data da baixa]                  datetime            NULL, 
       [Sequencia da entrada]           int                 DEFAULT 0 NOT NULL, 
       [Sequencia da baixa]             int                 DEFAULT 0 NOT NULL, 
       Documento                        varchar (25)        DEFAULT '' NOT NULL, 
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       [Sequencia do movimento]         int                 DEFAULT 0 NOT NULL, 
       Teste                            varchar (3)         DEFAULT '' NOT NULL, 
          CONSTRAINT [Sequencia da manutenção Pg] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Manutenção]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                  Manutenção contas receber                 */
/*------------------------------------------------------------*/

 CREATE TABLE [Manutenção contas receber] (
       [Sequencia da Manutenção]        int                 NOT NULL     IDENTITY (1,  1 ) NOT FOR REPLICATION, 
       Parcela                          smallint            DEFAULT 0 NOT NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       [Data de entrada]                datetime            NULL, 
       Historico                        text                DEFAULT '' NOT NULL, 
       [Especie de pagamento]           varchar (14)        DEFAULT '' NOT NULL, 
       [Forma de pagamento]             varchar (5)         DEFAULT '' NOT NULL, 
       [Data de vencimento]             datetime            NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Valor da parcela]               decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Valor pago]                     decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Valor do juros]                 decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Valor do desconto]              decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor restante]                 decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Data da baixa]                  datetime            NULL, 
       [Sequencia da entrada]           int                 DEFAULT 0 NOT NULL, 
       [Sequencia da baixa]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       Documento                        varchar (25)        DEFAULT '' NOT NULL, 
       [Sequencia da devolução]         int                 DEFAULT 0 NOT NULL, 
       [Numero da nfe]                  int                 DEFAULT 0 NOT NULL, 
       [Sequencia da saida]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia da Ordem]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia da Cobrança]          smallint            DEFAULT 0 NOT NULL, 
       [Sequencia da Venda]             int                 DEFAULT 0 NOT NULL, 
       [Tipo de conta]                  varchar (40)        DEFAULT '' NOT NULL, 
       Teste                            varchar (3)         DEFAULT '' NOT NULL, 
          CONSTRAINT Seq_ManuReceber PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Manutenção]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                           Marcas                           */
/*------------------------------------------------------------*/

 CREATE TABLE Marcas (
       [Sequencia da Marca]             int                 DEFAULT 0 NOT NULL, 
       Descrição                        varchar (80)        DEFAULT '' NOT NULL, 
          CONSTRAINT [Sequencia da Marca] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Marca]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                    Modelo do Virabrequim                   */
/*------------------------------------------------------------*/

 CREATE TABLE [Modelo do Virabrequim] (
       [Sequencia do Modelo]            smallint            DEFAULT 0 NOT NULL, 
       [Descrição do Modelo]            varchar (80)        DEFAULT '' NOT NULL, 
       [Sequencia da Marca]             int                 DEFAULT 0 NOT NULL, 
       [Peso do Virabrequim]            decimal (7, 3)      DEFAULT 0 NOT NULL, 
       [Linha do Modelo]                varchar (6)         DEFAULT '' NOT NULL, 
          CONSTRAINT [Sequencia do Modelo] PRIMARY KEY CLUSTERED
          (
                  [Sequencia do Modelo]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                    Movimento do estoque                    */
/*------------------------------------------------------------*/

 CREATE TABLE [Movimento do estoque] (
       [Sequencia do movimento]         int                 NOT NULL     IDENTITY (1,  1 ) NOT FOR REPLICATION, 
       Documento                        varchar (25)        DEFAULT '' NOT NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       [Data de entrada]                datetime            NULL, 
       [Forma de pagamento]             varchar (5)         DEFAULT '' NOT NULL, 
       Historico                        text                DEFAULT '' NOT NULL, 
       Cancelado                        bit                 DEFAULT 0 NOT NULL, 
       [Valor do frete]                 decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor do desconto]              decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total dos produtos]       decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total do ipi]             decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total do movimento]       decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Valor total icms st]            decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       [Data da Alteração]              datetime            NULL, 
       [Hora da Alteração]              datetime            NULL, 
       [Usuario da Alteração]           varchar (30)        DEFAULT '' NOT NULL, 
       Fechado                          bit                 DEFAULT 0 NOT NULL, 
       [Entrada Conserto]               bit                 DEFAULT 0 NOT NULL, 
       [Tipo de conta]                  varchar (40)        DEFAULT '' NOT NULL, 
          CONSTRAINT [Sequencia do movimento] PRIMARY KEY CLUSTERED
          (
                  [Sequencia do movimento]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                         Municipios                         */
/*------------------------------------------------------------*/

 CREATE TABLE Municipios (
       [Sequencia do municipio]         int                 DEFAULT 0 NOT NULL, 
       [Descrição do municipio]         varchar (80)        DEFAULT '' NOT NULL, 
       Uf                               varchar (3)         DEFAULT '' NOT NULL, 
       [Codigo do ibge]                 int                 DEFAULT 0 NOT NULL, 
       Cep                              varchar (9)         DEFAULT '' NOT NULL, 
       Inativo                          bit                 DEFAULT 0 NOT NULL, 
          CONSTRAINT [Sequencia do municipio] PRIMARY KEY CLUSTERED
          (
                  [Sequencia do municipio]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                    Natureza de operação                    */
/*------------------------------------------------------------*/

 CREATE TABLE [Natureza de operação] (
       [Sequencia da natureza]          smallint            DEFAULT 0 NOT NULL, 
       Cfop                             int                 DEFAULT 0 NOT NULL, 
       [Descrição do cfop]              varchar (40)        DEFAULT '' NOT NULL, 
       Inativo                          bit                 DEFAULT 0 NOT NULL, 
          CONSTRAINT [Sequencia da natureza] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da natureza]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                        NFe devolução                       */
/*------------------------------------------------------------*/

 CREATE TABLE [NFe devolução] (
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       [Chave da nfe]                   varchar (44)        DEFAULT '' NOT NULL, 
          CONSTRAINT [SeqNFe Devolucao] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da nota fiscal], [Chave da nfe]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                         Nota fiscal                        */
/*------------------------------------------------------------*/

 CREATE TABLE [Nota fiscal] (
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       [Numero da nfe]                  int                 DEFAULT 0 NOT NULL, 
       [Data de emissão]                datetime            NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia da natureza]          smallint            DEFAULT 0 NOT NULL, 
       [Sequencia da transportadora]    int                 DEFAULT 0 NOT NULL, 
       [Placa do veiculo]               varchar (8)         DEFAULT '' NOT NULL, 
       [Uf do veiculo]                  varchar (3)         DEFAULT '' NOT NULL, 
       Frete                            varchar (12)        DEFAULT '' NOT NULL, 
       [Valor do frete]                 decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor do desconto]              decimal (11, 2)     DEFAULT 0 NOT NULL, 
       Volumes                          int                 DEFAULT 0 NOT NULL, 
       Especie                          varchar (30)        DEFAULT '' NOT NULL, 
       [Data de saida]                  datetime            NULL, 
       [Hora da saida]                  datetime            NULL, 
       [Forma de pagamento]             varchar (5)         DEFAULT '' NOT NULL, 
       Historico                        text                DEFAULT '' NOT NULL, 
       [Valor total do ipi]             decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total do icms]            decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total dos produtos]       decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total da nfe]             decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Tipo de nota]                   smallint            DEFAULT 0 NOT NULL, 
       [Sequencia da classificação]     int                 DEFAULT 0 NOT NULL, 
       [Nota cancelada]                 bit                 DEFAULT 0 NOT NULL, 
       [Nota avulsa]                    bit                 DEFAULT 0 NOT NULL, 
       [Peso bruto]                     decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Peso liquido]                   decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Nfe complementar]               bit                 DEFAULT 0 NOT NULL, 
       [Chave nfe referenciada]         varchar (45)        DEFAULT '' NOT NULL, 
       [Chave de acesso da nfe]         varchar (50)        DEFAULT '' NOT NULL, 
       [Protocolo de autorização]       varchar (50)        DEFAULT '' NOT NULL, 
       [Data e hora da nfe]             varchar (25)        DEFAULT '' NOT NULL, 
       Transmitido                      bit                 DEFAULT 0 NOT NULL, 
       Autorizada                       bit                 DEFAULT 0 NOT NULL, 
       [Numero do recibo]               varchar (20)        DEFAULT '' NOT NULL, 
       Observação                       text                DEFAULT '' NOT NULL, 
       [Valor total da bc]              decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total do pis]             decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total do cofins]          decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total base st]            decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total icms st]            decimal (11, 2)     DEFAULT 0 NOT NULL, 
       Imprimiu                         bit                 DEFAULT 0 NOT NULL, 
       [Valor total dos tributos]       decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Finalidade nfe]                 smallint            DEFAULT 0 NOT NULL, 
       Devolução                        bit                 DEFAULT 0 NOT NULL, 
       [Data da Alteração]              datetime            NULL, 
       [Hora da Alteração]              datetime            NULL, 
       [Usuario da Alteração]           varchar (30)        DEFAULT '' NOT NULL, 
       XmlAssinado                      text                DEFAULT '' NOT NULL, 
       XmlAutorizado                    text                DEFAULT '' NOT NULL, 
       [Sequencia da devolução]         int                 DEFAULT 0 NOT NULL, 
       [Sequencia do movimento]         int                 DEFAULT 0 NOT NULL, 
       Conserto                         bit                 DEFAULT 0 NOT NULL, 
       [Sequencia da saida]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia da Venda]             int                 DEFAULT 0 NOT NULL, 
       [Valor Total dos Serviços]       decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor do Iss]                   decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Qual Empresa]                   smallint            DEFAULT 0 NOT NULL, 
          CONSTRAINT [Sequencia da nota fiscal] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da nota fiscal]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                      Ordem de Serviço                      */
/*------------------------------------------------------------*/

 CREATE TABLE [Ordem de Serviço] (
       [Sequencia da Ordem]             int                 NOT NULL, 
       [Data da Ordem]                  datetime            NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia do Vendedor]          int                 DEFAULT 0 NOT NULL, 
       [Forma de pagamento]             varchar (5)         DEFAULT '' NOT NULL, 
       [Valor total dos produtos]       decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor Total dos Serviços]       decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor do desconto]              decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor do frete]                 decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor da OS]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Aliquota do Iss]                decimal (5, 2)      DEFAULT 0 NOT NULL, 
       [Valor do Iss]                   decimal (10, 2)     DEFAULT 0 NOT NULL, 
       Cancelado                        bit                 DEFAULT 0 NOT NULL, 
       Fechado                          bit                 DEFAULT 0 NOT NULL, 
       Obs                              text                DEFAULT '' NOT NULL, 
       [Usuario da Alteração]           varchar (30)        DEFAULT '' NOT NULL, 
       [Data da Alteração]              datetime            NULL, 
       [Hora da Alteração]              datetime            NULL, 
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       [Sequencia do Transportador]     int                 DEFAULT 0 NOT NULL, 
       [Data de Cancelamento]           datetime            NULL, 
       [Qual Empresa]                   smallint            DEFAULT 0 NOT NULL, 
          CONSTRAINT [Sequencia da Ordem] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Ordem]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                    Outras saidas estoque                   */
/*------------------------------------------------------------*/

 CREATE TABLE [Outras saidas estoque] (
       [Sequencia da saida]             int                 NOT NULL     IDENTITY (1,  1 ) NOT FOR REPLICATION, 
       Documento                        varchar (25)        DEFAULT '' NOT NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       [Data de emissão]                datetime            NULL, 
       [Forma de pagamento]             varchar (5)         DEFAULT '' NOT NULL, 
       Historico                        text                DEFAULT '' NOT NULL, 
       Cancelado                        bit                 DEFAULT 0 NOT NULL, 
       [Valor total dos produtos]       decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total do ipi]             decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total icms st]            decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       Fechado                          bit                 DEFAULT 0 NOT NULL, 
       [Valor do frete]                 decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor do desconto]              decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total do movimento]       decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Usuario da Alteração]           varchar (30)        DEFAULT '' NOT NULL, 
       [Data da Alteração]              datetime            NULL, 
       [Hora da Alteração]              datetime            NULL, 
       Teste                            varchar (3)         DEFAULT '' NOT NULL, 
          CONSTRAINT [Sequencia da saida] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da saida]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                      Parametros da nfe                     */
/*------------------------------------------------------------*/

 CREATE TABLE [Parametros da nfe] (
       [Codigo da uf]                   smallint            DEFAULT 0 NOT NULL, 
       [Serie da nfe]                   smallint            DEFAULT 0 NOT NULL, 
       Ambiente                         smallint            DEFAULT 0 NOT NULL, 
       [Horario de verao]               bit                 DEFAULT 0 NOT NULL, 
       [Aliquota simples nacional]      decimal (8, 4)      DEFAULT 0 NOT NULL, 
       [Diretorio 1 nfe homologação]    varchar (255)       DEFAULT '' NOT NULL, 
       [Diretorio 2 nfe homologação]    varchar (255)       DEFAULT '' NOT NULL, 
       [Diretorio 1 nfe produção]       varchar (255)       DEFAULT '' NOT NULL, 
       [Diretorio 2 nfe produção]       varchar (255)       DEFAULT '' NOT NULL, 
       [Certificado digital]            varchar (255)       DEFAULT '' NOT NULL, 
       [Simples nacional]               bit                 DEFAULT 0 NOT NULL, 
       [Percentual tributos]            decimal (8, 4)      DEFAULT 0 NOT NULL, 
       [Micro Empresa]                  bit                 DEFAULT 0 NOT NULL, 
       [Usa Resp Tecnico]               bit                 DEFAULT 0 NOT NULL, 
       [CNPJ do Contador]               varchar (18)        DEFAULT '' NOT NULL
 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                      Parametros da OS                      */
/*------------------------------------------------------------*/

 CREATE TABLE [Parametros da OS] (
       [Letra Inicial]                  varchar (1)         DEFAULT '' NOT NULL, 
       [Letra Final]                    varchar (1)         DEFAULT '' NOT NULL, 
       [Lote Inicial]                   int                 DEFAULT 0 NOT NULL
 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                       Parcelas da os                       */
/*------------------------------------------------------------*/

 CREATE TABLE [Parcelas da os] (
       [Sequencia da Ordem]             int                 DEFAULT 0 NOT NULL, 
       [Numero da parcela]              smallint            DEFAULT 0 NOT NULL, 
       Dias                             smallint            DEFAULT 0 NOT NULL, 
       [Data de vencimento]             datetime            NULL, 
       [Valor da parcela]               decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Especie de pagamento]           varchar (14)        DEFAULT '' NOT NULL, 
          CONSTRAINT Seq_Ordem_e_Pc PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Ordem], [Numero da parcela]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                      Parcelas da Venda                     */
/*------------------------------------------------------------*/

 CREATE TABLE [Parcelas da Venda] (
       [Sequencia da Venda]             int                 DEFAULT 0 NOT NULL, 
       [Numero da parcela]              smallint            DEFAULT 0 NOT NULL, 
       Dias                             smallint            DEFAULT 0 NOT NULL, 
       [Data de vencimento]             datetime            NULL, 
       [Valor da parcela]               decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Especie de pagamento]           varchar (14)        DEFAULT '' NOT NULL, 
          CONSTRAINT SeqVenda_Pc PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Venda], [Numero da parcela]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                     Parcelas dev compra                    */
/*------------------------------------------------------------*/

 CREATE TABLE [Parcelas dev compra] (
       [Sequencia da devolução]         int                 DEFAULT 0 NOT NULL, 
       [Numero da parcela]              smallint            DEFAULT 0 NOT NULL, 
       Dias                             smallint            DEFAULT 0 NOT NULL, 
       [Data de vencimento]             datetime            NULL, 
       [Valor da parcela]               decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Especie de pagamento]           varchar (14)        DEFAULT '' NOT NULL, 
          CONSTRAINT SeqDev_e_Pc PRIMARY KEY CLUSTERED
          (
                  [Sequencia da devolução], [Numero da parcela]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                   Parcelas Entradas Pagar                  */
/*------------------------------------------------------------*/

 CREATE TABLE [Parcelas Entradas Pagar] (
       [Sequencia da entrada]           int                 DEFAULT 0 NOT NULL, 
       [Numero da parcela]              smallint            DEFAULT 0 NOT NULL, 
       Dias                             smallint            DEFAULT 0 NOT NULL, 
       [Data de vencimento]             datetime            NULL, 
       [Valor da parcela]               decimal (10, 2)     DEFAULT 0 NOT NULL, 
          CONSTRAINT SeqEntrada_e_Pc PRIMARY KEY CLUSTERED
          (
                  [Sequencia da entrada], [Numero da parcela]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                  Parcelas Entradas Receber                 */
/*------------------------------------------------------------*/

 CREATE TABLE [Parcelas Entradas Receber] (
       [Sequencia da entrada]           int                 DEFAULT 0 NOT NULL, 
       [Numero da parcela]              smallint            DEFAULT 0 NOT NULL, 
       Dias                             smallint            DEFAULT 0 NOT NULL, 
       [Data de vencimento]             datetime            NULL, 
       [Valor da parcela]               decimal (10, 2)     DEFAULT 0 NOT NULL, 
          CONSTRAINT Seq_e_Pc_Receber PRIMARY KEY CLUSTERED
          (
                  [Sequencia da entrada], [Numero da parcela]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                 Parcelas movimento estoque                 */
/*------------------------------------------------------------*/

 CREATE TABLE [Parcelas movimento estoque] (
       [Sequencia do movimento]         int                 DEFAULT 0 NOT NULL, 
       [Numero da parcela]              smallint            DEFAULT 0 NOT NULL, 
       Dias                             smallint            DEFAULT 0 NOT NULL, 
       [Data de vencimento]             datetime            NULL, 
       [Valor da parcela]               decimal (10, 2)     DEFAULT 0 NOT NULL, 
          CONSTRAINT SeqMov_Pc PRIMARY KEY CLUSTERED
          (
                  [Sequencia do movimento], [Numero da parcela]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                    Parcelas nota fiscal                    */
/*------------------------------------------------------------*/

 CREATE TABLE [Parcelas nota fiscal] (
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       [Numero da parcela]              smallint            DEFAULT 0 NOT NULL, 
       Dias                             smallint            DEFAULT 0 NOT NULL, 
       [Data de vencimento]             datetime            NULL, 
       [Valor da parcela]               decimal (10, 2)     DEFAULT 0 NOT NULL, 
          CONSTRAINT SeqNF_Pc PRIMARY KEY CLUSTERED
          (
                  [Sequencia da nota fiscal], [Numero da parcela]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                   Parcelas outras saidas                   */
/*------------------------------------------------------------*/

 CREATE TABLE [Parcelas outras saidas] (
       [Sequencia da saida]             int                 DEFAULT 0 NOT NULL, 
       [Numero da parcela]              smallint            DEFAULT 0 NOT NULL, 
       Dias                             smallint            DEFAULT 0 NOT NULL, 
       [Data de vencimento]             datetime            NULL, 
       [Valor da parcela]               decimal (10, 2)     DEFAULT 0 NOT NULL, 
       [Especie de pagamento]           varchar (14)        DEFAULT '' NOT NULL, 
          CONSTRAINT SeqSaida_Pc PRIMARY KEY CLUSTERED
          (
                  [Sequencia da saida], [Numero da parcela]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                          Produtos                          */
/*------------------------------------------------------------*/

 CREATE TABLE Produtos (
       [Sequencia do produto]           bigint              DEFAULT 0 NOT NULL, 
       Descrição                        varchar (80)        DEFAULT '' NOT NULL, 
       Estoque                          decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Sequencia da unidade]           smallint            DEFAULT 0 NOT NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Sequencia da classificação]     int                 DEFAULT 0 NOT NULL, 
       Usado                            bit                 DEFAULT 0 NOT NULL, 
       Inativo                          bit                 DEFAULT 0 NOT NULL, 
       [Material adquirido de terceiro] bit                 DEFAULT 0 NOT NULL, 
       Importado                        bit                 DEFAULT 0 NOT NULL, 
       [Ultima nfe]                     datetime            NULL, 
       [Ultima entrada]                 datetime            NULL, 
       [Ultimo fornecedor]              int                 DEFAULT 0 NOT NULL, 
       Despesa                          bit                 DEFAULT 0 NOT NULL, 
       [Valor unitario]                 decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor de custo]                 decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Ultima venda]                   datetime            NULL, 
       [Codigo de Barras]               varchar (13)        DEFAULT '' NOT NULL, 
          CONSTRAINT [Sequencia do produto] PRIMARY KEY CLUSTERED
          (
                  [Sequencia do produto]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                   Produtos da nota fiscal                  */
/*------------------------------------------------------------*/

 CREATE TABLE [Produtos da nota fiscal] (
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       [Sequencia produto nota fiscal]  int                 NOT NULL     IDENTITY (1,  1 ) NOT FOR REPLICATION, 
       [Sequencia do produto]           bigint              DEFAULT 0 NOT NULL, 
       Quantidade                       decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor unitario]                 decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Valor do ipi]                   decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor do icms]                  decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Aliquota do ipi]                decimal (5, 2)      DEFAULT 0 NOT NULL, 
       [Aliquota do icms]               decimal (5, 2)      DEFAULT 0 NOT NULL, 
       [Percentual da redução]          decimal (5, 2)      DEFAULT 0 NOT NULL, 
       [Valor da base de calculo]       decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor do pis]                   decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor do cofins]                decimal (11, 2)     DEFAULT 0 NOT NULL, 
       Iva                              decimal (7, 4)      DEFAULT 0 NOT NULL, 
       [Base de calculo st]             decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor icms st]                  decimal (11, 2)     DEFAULT 0 NOT NULL, 
       Cfop                             int                 DEFAULT 0 NOT NULL, 
       Cst                              smallint            DEFAULT 0 NOT NULL, 
       [Aliquota icms st]               decimal (5, 2)      DEFAULT 0 NOT NULL, 
       [Valor dos tributos]             decimal (11, 2)     DEFAULT 0 NOT NULL, 
          CONSTRAINT SeqNf_e_produto PRIMARY KEY CLUSTERED
          (
                  [Sequencia da nota fiscal], [Sequencia produto nota fiscal]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                      Produtos da Venda                     */
/*------------------------------------------------------------*/

 CREATE TABLE [Produtos da Venda] (
       [Sequencia da Venda]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia do item]              int                 DEFAULT 0 NOT NULL, 
       [Sequencia do produto]           bigint              DEFAULT 0 NOT NULL, 
       Quantidade                       decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor unitario]                 decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Valor do ipi]                   decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Aliquota do ipi]                decimal (5, 2)      DEFAULT 0 NOT NULL, 
       [Valor icms st]                  decimal (11, 2)     DEFAULT 0 NOT NULL, 
          CONSTRAINT Seq_Venda_e_Item PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Venda], [Sequencia do item]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                  Produtos devolução compra                 */
/*------------------------------------------------------------*/

 CREATE TABLE [Produtos devolução compra] (
       [Sequencia da devolução]         int                 DEFAULT 0 NOT NULL, 
       [Sequencia do item]              int                 DEFAULT 0 NOT NULL, 
       [Sequencia do produto]           bigint              DEFAULT 0 NOT NULL, 
       Quantidade                       decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor unitario]                 decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Valor do ipi]                   decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor icms st]                  decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Aliquota do ipi]                decimal (5, 2)      DEFAULT 0 NOT NULL, 
       [Sequencia do movimento]         int                 DEFAULT 0 NOT NULL, 
          CONSTRAINT Seq_dev_item PRIMARY KEY CLUSTERED
          (
                  [Sequencia da devolução], [Sequencia do item]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                 Produtos movimento estoque                 */
/*------------------------------------------------------------*/

 CREATE TABLE [Produtos movimento estoque] (
       [Sequencia do movimento]         int                 DEFAULT 0 NOT NULL, 
       [Sequencia do produto]           bigint              DEFAULT 0 NOT NULL, 
       Quantidade                       decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor unitario]                 decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Valor do ipi]                   decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor icms st]                  decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Aliquota do ipi]                decimal (5, 2)      DEFAULT 0 NOT NULL, 
       [Sequencia do item]              int                 DEFAULT 0 NOT NULL, 
          CONSTRAINT Seq_Mov_Prod PRIMARY KEY CLUSTERED
          (
                  [Sequencia do movimento], [Sequencia do item]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                   Produtos saidas estoque                  */
/*------------------------------------------------------------*/

 CREATE TABLE [Produtos saidas estoque] (
       [Sequencia da saida]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia do item]              int                 DEFAULT 0 NOT NULL, 
       [Sequencia do produto]           bigint              DEFAULT 0 NOT NULL, 
       Quantidade                       decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor unitario]                 decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Valor do ipi]                   decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor icms st]                  decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Aliquota do ipi]                decimal (5, 2)      DEFAULT 0 NOT NULL, 
          CONSTRAINT Seq_Saida_Item PRIMARY KEY CLUSTERED
          (
                  [Sequencia da saida], [Sequencia do item]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                          Serviços                          */
/*------------------------------------------------------------*/

 CREATE TABLE Serviços (
       [Sequencia do Serviço]           int                 DEFAULT 0 NOT NULL, 
       [Descrição do Serviço]           varchar (200)       DEFAULT '' NOT NULL, 
       [Valor do Serviço]               decimal (10, 2)     DEFAULT 0 NOT NULL, 
       Inativo                          bit                 DEFAULT 0 NOT NULL, 
          CONSTRAINT [Sequencia do Serviço] PRIMARY KEY CLUSTERED
          (
                  [Sequencia do Serviço]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                   Serviços da Nota Fiscal                  */
/*------------------------------------------------------------*/

 CREATE TABLE [Serviços da Nota Fiscal] (
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       [Sequencia do item]              int                 DEFAULT 0 NOT NULL, 
       [Sequencia do Serviço]           int                 DEFAULT 0 NOT NULL, 
       Quantidade                       decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor unitario]                 decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
          CONSTRAINT NFe_Serviço PRIMARY KEY CLUSTERED
          (
                  [Sequencia da nota fiscal], [Sequencia do item]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                       Serviços da OS                       */
/*------------------------------------------------------------*/

 CREATE TABLE [Serviços da OS] (
       [Sequencia da Ordem]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia do Serviço]           int                 DEFAULT 0 NOT NULL, 
       Quantidade                       decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor unitario]                 decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Sequencia do Virabrequim]       int                 DEFAULT 0 NOT NULL, 
       [Sequencia do item]              int                 DEFAULT 0 NOT NULL, 
          CONSTRAINT SeqOS_e_Servico PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Ordem], [Sequencia do item], [Sequencia do Serviço]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                      Serviços da Venda                     */
/*------------------------------------------------------------*/

 CREATE TABLE [Serviços da Venda] (
       [Sequencia da Venda]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia do item]              int                 DEFAULT 0 NOT NULL, 
       [Sequencia do Serviço]           int                 DEFAULT 0 NOT NULL, 
       Quantidade                       decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor unitario]                 decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor total]                    decimal (12, 2)     DEFAULT 0 NOT NULL, 
          CONSTRAINT Seq_venda_Item PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Venda], [Sequencia do item]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                      Tipo da Cobrança                      */
/*------------------------------------------------------------*/

 CREATE TABLE [Tipo da Cobrança] (
       [Sequencia da Cobrança]          smallint            DEFAULT 0 NOT NULL, 
       [Descrição da Cobrança]          varchar (45)        DEFAULT '' NOT NULL, 
          CONSTRAINT [Sequencia da Cobrança] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Cobrança]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                          Unidades                          */
/*------------------------------------------------------------*/

 CREATE TABLE Unidades (
       [Sequencia da unidade]           smallint            DEFAULT 0 NOT NULL, 
       Descrição                        varchar (80)        DEFAULT '' NOT NULL, 
       [Sigla da unidade]               varchar (20)        DEFAULT '' NOT NULL, 
       Inativo                          bit                 DEFAULT 0 NOT NULL, 
          CONSTRAINT [Sequencia da unidade] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da unidade]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                           Vendas                           */
/*------------------------------------------------------------*/

 CREATE TABLE Vendas (
       [Sequencia da Venda]             int                 DEFAULT 0 NOT NULL, 
       [Data da Venda]                  datetime            NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       Observações                      text                DEFAULT '' NOT NULL, 
       [Valor total do ipi]             decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor total dos produtos]       decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor Total dos Serviços]       decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor Total da Venda]           decimal (12, 2)     DEFAULT 0 NOT NULL, 
       [Valor total icms st]            decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Sequencia do Vendedor]          int                 DEFAULT 0 NOT NULL, 
       [Forma de pagamento]             varchar (5)         DEFAULT '' NOT NULL, 
       [Valor Total do ISS]             decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor do frete]                 decimal (11, 2)     DEFAULT 0 NOT NULL, 
       [Valor do desconto]              decimal (11, 2)     DEFAULT 0 NOT NULL, 
       Cancelado                        bit                 DEFAULT 0 NOT NULL, 
       Fechado                          bit                 DEFAULT 0 NOT NULL, 
       [Sequencia da transportadora]    int                 DEFAULT 0 NOT NULL, 
       [Data da Alteração]              datetime            NULL, 
       [Hora da Alteração]              datetime            NULL, 
       [Usuario da Alteração]           varchar (30)        DEFAULT '' NOT NULL, 
       [Sequencia da nota fiscal]       int                 DEFAULT 0 NOT NULL, 
       [Tipo de Venda]                  bit                 DEFAULT 0 NOT NULL, 
       Teste                            varchar (3)         DEFAULT '' NOT NULL, 
          CONSTRAINT [Sequencia da Venda] PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Venda]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                         Virabrequim                        */
/*------------------------------------------------------------*/

 CREATE TABLE Virabrequim (
       [Sequencia do Virabrequim]       int                 DEFAULT 0 NOT NULL, 
       Identificação                    varchar (9)         DEFAULT '' NOT NULL, 
       [Sequencia do geral]             int                 DEFAULT 0 NOT NULL, 
       [Ultima Os]                      datetime            NULL, 
       [Valor do Virabrequim]           decimal (10, 2)     DEFAULT 0 NOT NULL, 
       Modelo                           varchar (30)        DEFAULT '' NOT NULL, 
       [Sequencia da Marca]             int                 DEFAULT 0 NOT NULL, 
       Obs                              text                DEFAULT '' NOT NULL, 
       [Sequencia do Modelo]            smallint            DEFAULT 0 NOT NULL, 
       [Sequencia da Ordem]             int                 DEFAULT 0 NOT NULL, 
          CONSTRAINT [Sequencia do Virabrequim] PRIMARY KEY CLUSTERED
          (
                  [Sequencia do Virabrequim]
          ) WITH FILLFACTOR = 90

 )
GO

 CREATE UNIQUE INDEX ID_Virabrequim ON Virabrequim
       (Identificação) WITH FILLFACTOR = 90
GO




/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                      Virabrequim da OS                     */
/*------------------------------------------------------------*/

 CREATE TABLE [Virabrequim da OS] (
       [Sequencia da Ordem]             int                 DEFAULT 0 NOT NULL, 
       [Sequencia do Virabrequim]       int                 DEFAULT 0 NOT NULL, 
       Quantidade                       decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Valor unitario]                 decimal (11, 4)     DEFAULT 0 NOT NULL, 
       [Tem Casco]                      bit                 DEFAULT 0 NOT NULL, 
          CONSTRAINT Seq_Virabrequim PRIMARY KEY CLUSTERED
          (
                  [Sequencia da Ordem], [Sequencia do Virabrequim]
          ) WITH FILLFACTOR = 90

 )
GO


/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                          PW~Grupos                         */
/*------------------------------------------------------------*/

 CREATE TABLE [PW~Grupos] (
       [PW~Nome]                        varchar (100)       DEFAULT ''   COLLATE SQL_Latin1_General_CP1_CI_AS  NOT NULL,
          CONSTRAINT [PW~Nome] PRIMARY KEY  CLUSTERED
          (
                  [PW~Nome]
          ) WITH FILLFACTOR = 90
)
GO

/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                         PW~Usuarios                        */
/*------------------------------------------------------------*/

 CREATE TABLE [PW~Usuarios] (
       [PW~Grupo]                       varchar (100)       DEFAULT ''   COLLATE SQL_Latin1_General_CP1_CI_AS  NOT NULL,
       [PW~Nome]                        varchar (100)       DEFAULT ''   COLLATE SQL_Latin1_General_CP1_CI_AS  NOT NULL,
       [PW~Senha]                       varchar (100)       DEFAULT ''   COLLATE SQL_Latin1_General_CP1_CI_AS  NOT NULL,
       [PW~Obs]                         varchar (100)       DEFAULT ''   COLLATE SQL_Latin1_General_CP1_CI_AS  NULL,
       CONSTRAINT [Chave usuario] PRIMARY KEY  CLUSTERED
          (
                  [PW~Nome], [PW~Senha]
          ) WITH FILLFACTOR = 90
)
GO

 CREATE INDEX [PW~Grupo] ON [PW~Usuarios]
        ([PW~Grupo]) WITH FILLFACTOR = 90
GO

/*------------------------------------------------------------*/
/*     Criação de Tabelas, Indices e Atribuição de Default    */
/*                         PW~Tabelas                         */
/*------------------------------------------------------------*/

 CREATE TABLE [PW~Tabelas] (
       [PW~Projeto]                     varchar (10)        DEFAULT ''   COLLATE SQL_Latin1_General_CP1_CI_AS  NOT NULL,
       [PW~Grupo]                       varchar (100)       DEFAULT ''   COLLATE SQL_Latin1_General_CP1_CI_AS  NOT NULL,
       [PW~Nome]                        varchar (100)       DEFAULT ''   COLLATE SQL_Latin1_General_CP1_CI_AS  NOT NULL,
       [PW~Permissoes]                  varchar (100)       DEFAULT ''   COLLATE SQL_Latin1_General_CP1_CI_AS  NOT NULL,
          CONSTRAINT [Chave tabelas] PRIMARY KEY  CLUSTERED
          (
                  [PW~Projeto], [PW~Grupo], [PW~Nome]
          ) WITH FILLFACTOR = 90
)
GO

/*------------------------------------------------------------*/
/*         Criação de triggers/foreign key constraints        */
/*------------------------------------------------------------*/

 ALTER TABLE [dbo].[Baixa contas pagar] ADD CONSTRAINT [FK_Baixa_contas_pagar_Manutenção_contas_pagar]
      FOREIGN KEY
         (  [Sequencia da Manutenção]  )
      REFERENCES [Manutenção contas pagar]
         (  [Sequencia da Manutenção]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Baixa contas receber] ADD CONSTRAINT [FK_Baixa_contas_receber_Manutenção_contas_receber]
      FOREIGN KEY
         (  [Sequencia da Manutenção]  )
      REFERENCES [Manutenção contas receber]
         (  [Sequencia da Manutenção]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Baixa contas receber] ADD CONSTRAINT [FK_Baixa_contas_receber_Tipo_da_Cobrança]
      FOREIGN KEY
         (  [Sequencia da Cobrança]  )
      REFERENCES [Tipo da Cobrança]
         (  [Sequencia da Cobrança]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Cheques Pre] ADD CONSTRAINT [FK_Cheques_Pre_Bancos]
      FOREIGN KEY
         (  [Codigo do Banco]  )
      REFERENCES Bancos
         (  [Codigo do Banco]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Cheques Pre] ADD CONSTRAINT [FK_Cheques_Pre_Geral]
      FOREIGN KEY
         (  [Sequencia do geral]  )
      REFERENCES Geral
         (  [Sequencia do geral]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Entrada Contas Pagar] ADD CONSTRAINT [FK_Entrada_Contas_Pagar_Geral]
      FOREIGN KEY
         (  [Sequencia do geral]  )
      REFERENCES Geral
         (  [Sequencia do geral]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Entrada Contas Receber] ADD CONSTRAINT [FK_Entrada_Contas_Receber_Geral]
      FOREIGN KEY
         (  [Sequencia do geral]  )
      REFERENCES Geral
         (  [Sequencia do geral]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Geral] ADD CONSTRAINT [FK_Geral_Municipios]
      FOREIGN KEY
         (  [Sequencia do municipio]  )
      REFERENCES Municipios
         (  [Sequencia do municipio]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Itens da OS] ADD CONSTRAINT [FK_Itens_da_OS_Produtos]
      FOREIGN KEY
         (  [Sequencia do produto]  )
      REFERENCES Produtos
         (  [Sequencia do produto]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Manutenção contas pagar] ADD CONSTRAINT [FK_Manutenção_contas_pagar_Geral]
      FOREIGN KEY
         (  [Sequencia do geral]  )
      REFERENCES Geral
         (  [Sequencia do geral]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Manutenção contas receber] ADD CONSTRAINT [FK_Manutenção_contas_receber_Geral]
      FOREIGN KEY
         (  [Sequencia do geral]  )
      REFERENCES Geral
         (  [Sequencia do geral]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Modelo do Virabrequim] ADD CONSTRAINT [FK_Modelo_do_Virabrequim_Marcas]
      FOREIGN KEY
         (  [Sequencia da Marca]  )
      REFERENCES Marcas
         (  [Sequencia da Marca]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Movimento do estoque] ADD CONSTRAINT [FK_Movimento_do_estoque_Geral]
      FOREIGN KEY
         (  [Sequencia do geral]  )
      REFERENCES Geral
         (  [Sequencia do geral]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Nota fiscal] ADD CONSTRAINT [FK_Nota_fiscal_Geral]
      FOREIGN KEY
         (  [Sequencia do geral]  )
      REFERENCES Geral
         (  [Sequencia do geral]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Nota fiscal] ADD CONSTRAINT [FK_Nota_fiscal_Natureza_de_operação]
      FOREIGN KEY
         (  [Sequencia da natureza]  )
      REFERENCES [Natureza de operação]
         (  [Sequencia da natureza]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Nota fiscal] ADD CONSTRAINT [FK_Nota_fiscal_Classificação_fiscal]
      FOREIGN KEY
         (  [Sequencia da classificação]  )
      REFERENCES [Classificação fiscal]
         (  [Sequencia da classificação]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Ordem de Serviço] ADD CONSTRAINT [FK_Ordem_de_Serviço_Geral]
      FOREIGN KEY
         (  [Sequencia do geral]  )
      REFERENCES Geral
         (  [Sequencia do geral]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Outras saidas estoque] ADD CONSTRAINT [FK_Outras_saidas_estoque_Geral]
      FOREIGN KEY
         (  [Sequencia do geral]  )
      REFERENCES Geral
         (  [Sequencia do geral]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Produtos] ADD CONSTRAINT [FK_Produtos_Classificação_fiscal]
      FOREIGN KEY
         (  [Sequencia da classificação]  )
      REFERENCES [Classificação fiscal]
         (  [Sequencia da classificação]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Produtos] ADD CONSTRAINT [FK_Produtos_Unidades]
      FOREIGN KEY
         (  [Sequencia da unidade]  )
      REFERENCES Unidades
         (  [Sequencia da unidade]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Produtos da nota fiscal] ADD CONSTRAINT [FK_Produtos_da_nota_fiscal_Produtos]
      FOREIGN KEY
         (  [Sequencia do produto]  )
      REFERENCES Produtos
         (  [Sequencia do produto]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Produtos da Venda] ADD CONSTRAINT [FK_Produtos_da_Venda_Produtos]
      FOREIGN KEY
         (  [Sequencia do produto]  )
      REFERENCES Produtos
         (  [Sequencia do produto]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Produtos devolução compra] ADD CONSTRAINT [FK_Produtos_devolução_compra_Produtos]
      FOREIGN KEY
         (  [Sequencia do produto]  )
      REFERENCES Produtos
         (  [Sequencia do produto]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Produtos movimento estoque] ADD CONSTRAINT [FK_Produtos_movimento_estoque_Produtos]
      FOREIGN KEY
         (  [Sequencia do produto]  )
      REFERENCES Produtos
         (  [Sequencia do produto]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Serviços da Nota Fiscal] ADD CONSTRAINT [FK_Serviços_da_Nota_Fiscal_Serviços]
      FOREIGN KEY
         (  [Sequencia do Serviço]  )
      REFERENCES Serviços
         (  [Sequencia do Serviço]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Serviços da OS] ADD CONSTRAINT [FK_Serviços_da_OS_Serviços]
      FOREIGN KEY
         (  [Sequencia do Serviço]  )
      REFERENCES Serviços
         (  [Sequencia do Serviço]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Serviços da Venda] ADD CONSTRAINT [FK_Serviços_da_Venda_Serviços]
      FOREIGN KEY
         (  [Sequencia do Serviço]  )
      REFERENCES Serviços
         (  [Sequencia do Serviço]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Vendas] ADD CONSTRAINT [FK_Vendas_Geral]
      FOREIGN KEY
         (  [Sequencia do geral]  )
      REFERENCES Geral
         (  [Sequencia do geral]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Virabrequim] ADD CONSTRAINT [FK_Virabrequim_Geral]
      FOREIGN KEY
         (  [Sequencia do geral]  )
      REFERENCES Geral
         (  [Sequencia do geral]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Virabrequim] ADD CONSTRAINT [FK_Virabrequim_Marcas]
      FOREIGN KEY
         (  [Sequencia da Marca]  )
      REFERENCES Marcas
         (  [Sequencia da Marca]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[Virabrequim da OS] ADD CONSTRAINT [FK_Virabrequim_da_OS_Virabrequim]
      FOREIGN KEY
         (  [Sequencia do Virabrequim]  )
      REFERENCES Virabrequim
         (  [Sequencia do Virabrequim]  )
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[PW~Usuarios] ADD CONSTRAINT [FK_PW~Usuarios_PW~Grupos]
      FOREIGN KEY
         ( [PW~Grupo])
      REFERENCES [PW~Grupos]
         ( [PW~Nome] )
      ON DELETE CASCADE
      ON UPDATE CASCADE
GO

 ALTER TABLE [dbo].[PW~Tabelas] ADD CONSTRAINT [FK_PW~Tabelas_PW~Grupos]
      FOREIGN KEY
         ( [PW~Grupo])
      REFERENCES [PW~Grupos]
         ( [PW~Nome] )
      ON DELETE CASCADE
      ON UPDATE CASCADE
GO
