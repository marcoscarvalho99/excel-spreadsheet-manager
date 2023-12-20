package com.example.gestorfinanciavendaexcel;

import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@AllArgsConstructor
@NoArgsConstructor
@Data
@Entity
public class Produto {
    @Id
    @GeneratedValue(strategy = GenerationType.AUTO)
    private Long id;
    private String nomeProduto;
    private String descricao;
    private String categoria;
    private String codigoProduto;
    private Double peso;
    private String dimensao;
    private Double preco;
    private int  quantEstoque;
    private String dataValidade;
    private String cor;
    private String tamanho;
    private String material;
    private String fabricante;
    private String paisOrigem;
    private String observacoes;
    private String codigobarra;
    private String localizacaoArmazem;




}
