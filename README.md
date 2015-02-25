# POIUtils

##Libreria de apoyo para usar Apache POI

Para poder usar la clase para leer archivos de excel, necesitamos la direccion del archivo y las cabeceras de las que nos interesa recuperar la informacion
```java
//Direccion del archivo que se va procesar
String pathFile = "/home/solrac/pruebas.xls";

//Cabeceras en el archivo que nos interesa procesar
String[] cabecera_text = "Código del servicio,Nombre del servicio,Impuesto 1,Precio 1".split(",");
```

**Opcional: Podemos hacer todas las cabeceras opcionales para recuperar toda la informacion aunque alguna celda este vacia
```java
//Hacemos a todas las cabeceras opcionales
List cabecera = new ArrayList();
for (String text : cabecera_text) {
    cabecera.add(new Header(text, false));
}
```
Creamos el objeto ExcelReader y recuperamos la informacion del archivo.
El metodo `getContent();` regresa la informacion en forma de List<HashMap>
```java
ExcelReader reader = new ExcelReader(pathFile, cabecera);
List<HashMap> hojas = reader.getContent();
```
Cada HashMap contiene 2 keys ("sheet_name" y "data").
  - sheet_name: nombre de la hoja
  - data: informacion de la hoja, cada fila es un `String[]`
```java
//Imprimir filas recuperadas
List<String[]> filas = (List<String[]>) hojas.get(0).get("data");
for (int j = 0; j < filas.size(); j++) {
    System.out.println("Fila " + (j + 1) + ": " + arrayToString(filas.get(j)));
}

//metodo arrayToString()
private String arrayToString(String[] info) {
    String cadena = "";
    for (int i = 0; i < info.length; i++) {
        cadena += info[i] + ",";
    }
    return cadena + "|";
}
```
