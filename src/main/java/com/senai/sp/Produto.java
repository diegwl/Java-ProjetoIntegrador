package com.senai.sp;

// Classe dos produtos da loja, com nome e preço.
public class Produto {
    private String nome;
    private double preco;

    public Produto(){}

    public Produto(String nome, double preco) {
        super();
        this.nome = nome;
        this.preco = preco;
    }

    public String getNome() {
        return nome;
    }

    public double getPreco() {
        return preco;
    }

    public void setNome(String nome) {
        this.nome = nome;
    }

    public void setPreco(double preco) {
        this.preco = preco;
    }
}
