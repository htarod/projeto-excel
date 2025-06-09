# Controle de Estoque em Excel com VBA

Este projeto implementa um sistema simples de controle de estoque automatizado utilizando Excel e VBA.
## Funcionalidades

- Formulário com interface amigável para registrar movimentações (entradas ou saídas)
- Lançamento automático em abas separadas: **Estoque**, **Entradas**, **Saídas** e **Log**
- Atualização automática da quantidade em estoque
- Cálculo de alerta de estoque mínimo com fórmula personalizada

## Estrutura da Planilha

- `Início`: página de boas-vindas com instruções e botão para abrir o formulário
- `Estoque`: lista atualizada dos produtos com cálculo de alerta automático
- `Entradas`: histórico das entradas registradas
- `Saídas`: histórico das saídas registradas
- `Log`: registro cronológico de todas as movimentações

## Fórmula de Alerta

Na aba **Estoque**, a seguinte fórmula é usada para exibir alertas quando a quantidade está abaixo do mínimo definido:

```excel
=SE(D2<E2; "Estoque Baixo"; "Normal")
