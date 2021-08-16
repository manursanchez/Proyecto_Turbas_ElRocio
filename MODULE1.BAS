Attribute VB_Name = "Module1"
'Declaracion de variables globales al programa
Public iva As Integer 'Variable que va a tener el
'valor del IVA en este caso 16
'Sitios donde se utiliza esta variable:
'--Formulario 6 en la primera carga de este

Public valor1euro As Currency 'Almacenara el Valor del euro
Public € As Currency
Public ptas As Currency
Public €iva As Currency
Public ptasiva As Currency
'Estas variables son de tipo moneda y las iremos
'utilizando a lo largo del programa para los materiales
Public numfactura As String
'Esta variable publica va a contener el numero de factura
'digamos que va a ser una variable auxiliar
'La utilizare cuando tenga que hacer busqueda de facturas
'o para traspasar de un formulario a otro el numero de factura
Public identificacion As String
'Esta variable public va a realizar la misma funcion
'que la variable anterior(numfactura) solamente que va a ser
'una variable auxiliar para los clientes

'Declaracion de variables para realizar los totales de las facturas
Public recargo As Currency
Public totalrecargo As Currency
Public ivas As Currency
Public totalivas As Currency
Public ptotal As Currency
Public sumatotal As Currency

'Declaracion de la variable global que almacenara el codigo de presupuesto
Public npres As String

