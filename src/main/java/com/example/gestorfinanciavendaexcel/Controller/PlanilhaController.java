package com.example.gestorfinanciavendaexcel.Controller;

import com.example.gestorfinanciavendaexcel.Produto;
import com.example.gestorfinanciavendaexcel.ProdutoRepository;
import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import org.springframework.web.servlet.view.RedirectView;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;


@Controller
@RequestMapping("planilha")
public class PlanilhaController {

    @Autowired
    ProdutoRepository produtoRepository;

    @GetMapping("/send-form")
    public ModelAndView enviarPlanilha(){

        ModelAndView model = new ModelAndView("uploadPlanilha");
        return model;
    }
    @GetMapping("/redirecionar")
    public RedirectView redirecionar(RedirectAttributes attributes) {
        // Adicionar mensagem ao Flash Attributes
        attributes.addFlashAttribute("mensagem", "Redirecionamento bem-sucedido!");

        // Redirecionar para a URL "/outraPagina"
        return new RedirectView("/uploadPlanilha");
    }

    @PostMapping("/upload")
    public ModelAndView uploadExcelFile(@RequestParam("file") MultipartFile file, RedirectAttributes redirectAttributes) throws IOException {
        ModelAndView modelAndView = new ModelAndView("/uploadPlanilha");

        Workbook workbook = WorkbookFactory.create(file.getInputStream());

        if (!file.isEmpty()) {
            // Supondo que seu arquivo Excel tem apenas uma planilha
            Sheet sheet = workbook.getSheetAt(0);

//                // Convertendo a planilha para um formato mais fácil de manipular (lista de listas)
            List<List<String>> excelData = new ArrayList<>();
            for (Row row : sheet) {
                List<String> rowData = new ArrayList<>();
                for (Cell cell : row) {
                    rowData.add(cell.toString());
                }
                excelData.add(rowData);
            }
            AtomicBoolean flag = new AtomicBoolean(false);
            excelData.stream().forEach( produto ->{
                System.out.println(produto.toString());
                if (flag.get()){
                    Produto p = new Produto();
                    p.setNomeProduto(produto.get(0).toString());
                    p.setDescricao(produto.get(1).toString());
                    p.setCategoria(produto.get(2).toString());
                    p.setCodigoProduto(produto.get(3).toString());
                    p.setPeso(Double.parseDouble(produto.get(4)));
                    p.setDimensao(produto.get(5).toString());
                    p.setPreco(Double.parseDouble(produto.get(6)));
                   String estoque =  produto.get(7);
                    p.setQuantEstoque(Integer.parseInt(estoque.replace(".","")));
                    p.setDataValidade(produto.get(8));
                    p.setCor(produto.get(9));
                    p.setTamanho(produto.get(10));
                    p.setMaterial(produto.get(11));
                    p.setFabricante(produto.get(12));
                    p.setPaisOrigem(produto.get(13));
                    p.setObservacoes(produto.get(14));
                    p.setCodigobarra(produto.get(15));
                    p.setLocalizacaoArmazem(produto.get(16));

                     produtoRepository.save(p);
                }

                flag.set(true);
            });




//                modelAndView.addObject("excelData", excelData);
        } else {
            modelAndView.addObject("error", "Por favor, selecione um arquivo Excel.");
        }
        modelAndView.addObject("message","adicionado com sucesso!");
//        redirectAttributes.addFlashAttribute("message","adicionado com sucesso!");
        return modelAndView;
    }
//    http://localhost:8080/planilha/download
    @GetMapping("/download")
    public ResponseEntity<InputStreamResource> downloadExcel() {
       List<Produto> produtos = produtoRepository.findAll();
        List<String> lista = List.of("Nome", "Descricao", "Categoria", "Codigo" , "Peso"
                , "Dimensão", "Preço", "Estoque" , "Validade", "cor", "tamanho", "Material", "Fabricante", "Pais de origem", "Observações", "Codigo de Barras", "Localização Armazem");

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Planilha");

            int rowNum = 0;
            for (String item : lista) {
                Row row = sheet.createRow(rowNum++);
                Cell cell = row.createCell(0);
                cell.setCellValue("Nome");
                cell = row.createCell(1);
                cell.setCellValue("Descricao");
                cell = row.createCell(2);
                cell.setCellValue("Categoria");
                cell = row.createCell(3);
                cell.setCellValue("Codigo");
                cell = row.createCell(4);
                cell.setCellValue("Peso");
                cell = row.createCell(5);
                cell.setCellValue("Dimensão");
                cell = row.createCell(6);
                cell.setCellValue("Preço");
                cell = row.createCell(7);
                cell.setCellValue("Estoque");
                cell = row.createCell(8);
                cell.setCellValue("Validade");
                cell = row.createCell(9);
                cell.setCellValue("cor");
                cell = row.createCell(10);
                cell.setCellValue("tamanho");
                cell = row.createCell(11);
                cell.setCellValue("Material");
                cell = row.createCell(12);
                cell.setCellValue("Fabricante");
                cell = row.createCell(13);
                cell.setCellValue("Pais de origem");
                cell = row.createCell(14);
                cell.setCellValue("Observações");
                cell = row.createCell(15);
                cell.setCellValue("Codigo de Barras");
                cell = row.createCell(16);
                cell.setCellValue("Localização Armazem");

            }
            rowNum = 1;
            for (Produto item : produtos) {
                Row row = sheet.createRow(rowNum++);
                Cell cell = row.createCell(0);
                cell.setCellValue(item.getNomeProduto());
                cell = row.createCell(1);
                cell.setCellValue(item.getDescricao());
                cell = row.createCell(2);
                cell.setCellValue(item.getCategoria());
                cell = row.createCell(3);
                cell.setCellValue(item.getCodigoProduto());
                cell = row.createCell(4);
                cell.setCellValue(item.getPeso());
                cell = row.createCell(5);
                cell.setCellValue(item.getDimensao());
                cell = row.createCell(6);
                cell.setCellValue(item.getPreco());
                cell = row.createCell(7);
                cell.setCellValue(item.getQuantEstoque());
                cell = row.createCell(8);
                cell.setCellValue(item.getDataValidade());
                cell = row.createCell(9);
                cell.setCellValue(item.getCor());
                cell = row.createCell(10);
                cell.setCellValue(item.getTamanho());
                cell = row.createCell(11);
                cell.setCellValue(item.getMaterial());
                cell = row.createCell(12);
                cell.setCellValue(item.getFabricante());
                cell = row.createCell(13);
                cell.setCellValue(item.getPaisOrigem());
                cell = row.createCell(14);
                cell.setCellValue(item.getObservacoes());
                cell = row.createCell(15);
                cell.setCellValue(item.getCodigobarra());
                cell = row.createCell(16);
                cell.setCellValue(item.getLocalizacaoArmazem());

            }

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            workbook.write(bos);


            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=planilha.xlsx");


            return ResponseEntity
                    .ok()
                    .headers(headers)
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .body(new InputStreamResource(new ByteArrayInputStream(bos.toByteArray())));
        } catch (IOException e) {
            e.printStackTrace();
            return ResponseEntity
                    .status(500)
                    .body(null);
        }
    }
}
