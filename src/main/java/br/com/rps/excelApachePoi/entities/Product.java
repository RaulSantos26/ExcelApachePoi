package br.com.rps.excelApachePoi.entities;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Product {
    private Number N;
    private Number CD_PRC;
    private Number CD_PRF_UND;
    private String CD_EQP;

    private Number CD_CLI;


    public Product(int n, long cdPrc, long prf, long cli) {
    }
}
