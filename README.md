# Amostra de funcionalidade: Impressao Inteligente!

## Funcionalidade de Impressão.

### classe principal: cPrint.cls
 
> Entendendo a classe principal é possível entender por extensão toda a estrutura.
 
 Do ponto de vista da arquitetura de projeto MVC, este é o `Controller` principal da
 funcionalidade de impressão. (o `View` é a própria folha impressa, e o `Model` é a classe
 cPedido_Agregado - obtém, processa e disponibiliza dados para a funcionalidade impressão -.

 Do ponto de vista da estrutura interna da funcionalidade de impressão, esta classe é
 o `Context` do pattern `strategy`.

 Do ponto de vista da relação desta funcionalidade com o resto do sistema, esta classe
 faz o papel de `fachada` que atende todos os pedidos de impressão.

 Responsabilidade: receber os mais diversos pedidos de impressão e direcionar aos
 mecanismos internos que atendam a solicitação.

### Propósito Didático e Ilustrativo

Aqui são apresentadas classes e módulos que integram a funcionalidade de impressão,
as classes preservam todas as características técnicas originais, mas não apresentam
todas as regras de negócio nem todos os métodos para preservar detalhes de implementação.
