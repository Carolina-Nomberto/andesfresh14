package com.andesfresh.portal.controller;

import com.andesfresh.portal.model.Empleado;
import com.andesfresh.portal.model.Producto;
import com.andesfresh.portal.service.EmpleadoService;
import com.andesfresh.portal.service.ProductoService;
import jakarta.servlet.http.HttpServletResponse;
import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

@Controller
@RequestMapping("/reportes")
public class ReporteController {

    private final ProductoService productoService;
    private final EmpleadoService empleadoService;

    public ReporteController(ProductoService productoService, EmpleadoService empleadoService) {
        this.productoService = productoService;
        this.empleadoService = empleadoService;
    }

    // ---------- 1) CSV ----------
    @GetMapping("/csv")
    public void exportarCsv(HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setCharacterEncoding("UTF-8");
        response.setHeader("Content-Disposition", "attachment; filename=productos.csv");

        PrintWriter writer = response.getWriter();
        writer.write("\ufeff"); // BOM: Excel reconoce las tildes
        writer.println("ID,Nombre,Marca,Categoria,Precio,Cantidad");
        for (Producto p : productoService.listar()) {
            writer.printf(Locale.US, "%d,%s,%s,%s,%.2f,%d%n",
                    p.getId(), p.getNombre(), p.getMarca(), p.getCategoria(),
                    p.getPrecio(), p.getCantidad());
        }
        writer.flush();
    }

    // ---------- 2) Excel con Apache POI ----------
    @GetMapping("/excel")
    public void exportarExcel(HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=productos.xlsx");

        List<Producto> productos = productoService.listar();
        try (Workbook libro = new XSSFWorkbook()) {
            Sheet hoja = libro.createSheet("Productos");

            // Estilo de la cabecera: negrita
            Font negrita = libro.createFont();
            negrita.setBold(true);
            CellStyle estiloCabecera = libro.createCellStyle();
            estiloCabecera.setFont(negrita);

            String[] cabecera = {"ID", "Nombre", "Marca", "Categoría", "Precio", "Cantidad", "Subtotal"};
            Row fila0 = hoja.createRow(0);
            for (int i = 0; i < cabecera.length; i++) {
                fila0.createCell(i).setCellValue(cabecera[i]);
                fila0.getCell(i).setCellStyle(estiloCabecera);
                hoja.setColumnWidth(i, 18 * 256); // ancho fijo (18 caracteres)
            }

            int numFila = 1;
            for (Producto p : productos) {
                Row fila = hoja.createRow(numFila);
                int filaExcel = numFila + 1; // en Excel las filas empiezan en 1
                fila.createCell(0).setCellValue(p.getId());
                fila.createCell(1).setCellValue(p.getNombre());
                fila.createCell(2).setCellValue(p.getMarca());
                fila.createCell(3).setCellValue(p.getCategoria());
                fila.createCell(4).setCellValue(p.getPrecio());
                fila.createCell(5).setCellValue(p.getCantidad());
                fila.createCell(6).setCellFormula("E" + filaExcel + "*F" + filaExcel);
                numFila++;
            }

            // Fila de total con fórmula SUM
            Row total = hoja.createRow(numFila);
            total.createCell(5).setCellValue("TOTAL");
            total.getCell(5).setCellStyle(estiloCabecera);
            total.createCell(6).setCellFormula("SUM(G2:G" + numFila + ")");

            libro.write(response.getOutputStream());
        }
    }

    // ---------- 3) PDF con JasperReports ----------
    @GetMapping("/pdf")
    public void generarPdf(HttpServletResponse response) throws Exception {
        try (InputStream plantilla = getClass().getResourceAsStream("/reportes/productos.jrxml")) {
            JasperReport report = JasperCompileManager.compileReport(plantilla);

            JRBeanCollectionDataSource datos = new JRBeanCollectionDataSource(productoService.listar());
            Map<String, Object> parametros = new HashMap<>();
            parametros.put("titulo", "Catálogo de productos - Andes Fresh S.A.C.");

            JasperPrint jasperPrint = JasperFillManager.fillReport(report, parametros, datos);

            response.setContentType("application/pdf");
            response.setHeader("Content-Disposition", "inline; filename=productos.pdf");
            JasperExportManager.exportReportToPdfStream(jasperPrint, response.getOutputStream());
        }
    }

    // ---------- 4) Excel de Empleados (Reto Opcional) ----------
    @GetMapping("/empleados/excel")
    public void exportarEmpleadosExcel(HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=empleados.xlsx");

        List<Empleado> empleados = empleadoService.listar();
        DateTimeFormatter formatoFecha = DateTimeFormatter.ofPattern("dd/MM/yyyy");

        try (Workbook libro = new XSSFWorkbook()) {
            Sheet hoja = libro.createSheet("Empleados");

            Font negrita = libro.createFont();
            negrita.setBold(true);
            CellStyle estiloCabecera = libro.createCellStyle();
            estiloCabecera.setFont(negrita);

            String[] cabecera = {"ID", "Nombres", "Apellidos", "Fecha de ingreso", "Sueldo", "Dirección"};
            Row fila0 = hoja.createRow(0);
            for (int i = 0; i < cabecera.length; i++) {
                fila0.createCell(i).setCellValue(cabecera[i]);
                fila0.getCell(i).setCellStyle(estiloCabecera);
                hoja.setColumnWidth(i, 20 * 256);
            }

            int numFila = 1;
            for (Empleado e : empleados) {
                Row fila = hoja.createRow(numFila);
                fila.createCell(0).setCellValue(e.getId());
                fila.createCell(1).setCellValue(e.getNombres());
                fila.createCell(2).setCellValue(e.getApellidos());
                fila.createCell(3).setCellValue(e.getFechaIngreso() != null ? e.getFechaIngreso().format(formatoFecha) : "");
                fila.createCell(4).setCellValue(e.getSueldo());
                fila.createCell(5).setCellValue(e.getDireccion());
                numFila++;
            }

            // Fila de total con fórmula SUM en la columna E (Sueldo)
            Row total = hoja.createRow(numFila);
            total.createCell(3).setCellValue("TOTAL SUELDOS");
            total.getCell(3).setCellStyle(estiloCabecera);

            Cell celdaTotal = total.createCell(4);
            celdaTotal.setCellFormula("SUM(E2:E" + numFila + ")");
            celdaTotal.setCellStyle(estiloCabecera);

            libro.write(response.getOutputStream());
        }
    }
}