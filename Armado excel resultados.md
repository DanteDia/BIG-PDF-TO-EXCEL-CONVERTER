# Armado excel resultados  
  
Objetivo: Este merge tiene que nuclear la informacion de ambos archivos para tener en un solo excel todos los movimientos.  
El proceso es una traduccion del esquema de gallo hacia el esquema de visual  
  
Hojas principales: las mismas que visual + posicion inicial(gallo) + posicion final(gallo-para extraer tenencias de aca) + Rentas y dividendos Gallo (creada)  
  
Hojas Auxiliares: EspecieVisual(hoja mapping)+EspecieGallo(hoja mapping codigo de especie a nombre) + precios iniciales especies (gallo) + cotizacion dolar historica  
  
  
HOJA POSICION INICIAL (OK)  
Se extrae directamente de gallo la hoja de posicion inicial. Modificaciones:   
1. Si hay activos sin precio pero con cantidad (col e) e importe (col g) lo calculamos como importe/cantidad.   
2. Dividimos la columna especie en dos: la primer palabra se va a la nueva columna B: Ticker y el resto se queda en la columna C: Especie  
  
  
HOJA POSICION FINAL MERGE (OK, falta desarrollar precio costo)   
  
Se arma la hoja de posicion final Gallo que es el punto cero de nuestra base de datos de tenencias para extraer la info de gallo y volcar en visual. Esta base se actualiza con cada movimiento de visual. Se extraen del archivo de gallo la hoja ‘Posicion Final’ entera y se encuentra y completa con unas columnas:  
*  Primero modificamos la columna especie y la separamos en dos, la primer palabra se convierte en una nueva columna: ticker y el resto queda como especie. ej: pamp pampa holding se divide en col ticker (pamp) y col especie(pampa holding)   
* Codigo de especie: se extrae de las otras hojas. Aca tenes que buscar la palabra clave de la columna B (especie) de la hoja posicion final en las otras hojas para hacer el match con el cod_especie. Para encontrarlo cabe aclarar que el match no es perfecto por ejemplo en la posicion final puede aparecer KO CEDEAR COCA-COLA COMPANY con su tipo_especie (col) titulos privados locales (que es lo mismo que tituloa privados exentos) y cuando vamos a esa hoja del excel (tit. Privados exentos) no esta el KO inicial y dice solo CEDEAR COCA-COLA COMPANY …y completamos con 8006. En los casos que el nombre de la especie tiene una fecha dentro es porque ese es el vencimiento y debe estar incluida si o si en el match de nombres para asignarle el codigo. Si no se opero esa especie en el periodo es probable que no este y ahi se usa la hoja de EspeciesGallo para mappear y conseguir el codigo.   
* Columna Codigo Especie Origen: si el codigo se obtuvo directamente del archivo de gallo va ‘Gallo’ y si se obtuvo del mapping ‘EspeciesGallo’.  
* Columna comentarios especies: aca se tratan algunos casos especiales. Lee detenidamente porque afectan la logica de armado, algunos codigoa si no se encuentran o hay mas de uno se ve en visual cual aparece. Esta columna no va en el resultado final del merge.   
* Columna precio tenencia final pesos: columna importe_pesos/columna cantidad   
* Columna precio tenencia final usd: columna importe en dolares/columna cantidad.   
* Columna Precio Tenencia Inicial: extraemos el precio de la hoja PreciosInicialesEspecies a traves de un buscarv via el ticker. Aca hay tres valores que son fijos basandonos en el ticker: pesos es 1, dolares es 1167.806 y dolar cable es 1148.93 para la fecha 1/1 o 31/12   
* Precio a utilizar: aca por ahora usamos precio tenencia inicial. Mas tarde usaremos el precio costo si se puede calcular bien.   
* Precio costo: Columna precio costo: para las especies que tuvieron transacciones en el periodo del pdf veremos que aparecen todas las compras y ponderando cantidadxprecio podemos obtener el precio costo.  
    *  Si vemos el ejemplo de Aluar Aluar vemos en la hoja tit privados exentos que hay 6 compras en distintas fechas. Podemos notar que la posicion final nos indica que hay 844 cantidades (col h de posicion final) y esto coincide con la suma de las cantidades de operaciones compra de la hoja de tit privados exentos. Asi mismo sabemos que si sumamos la cantidad (col g) por el precio (col h) (de las operaciones de compra, dejamos la de dividendo de lado) y dividimos por la cantidad de la posicion final (844 en este ejemplo) obtenemos el precio costo promedio 959.12 de aluar. Esto solo se puede hacer para las especies que solo tienen operaciones de compra, venta, amortizaciones y dividendos . Si la especie tiene en el mix de operaciones canjes entonces pongo el precio de la posicion inicial. Las operaciones de compra y venta puede llamarse de distintas formas como: compra, venta Usd, compra usd, cpra cable, venta.   
    * Veamos otro ejemplo el de AL30 BONO USD 2030 LA si vemos en la hoja posicion final tiene 18400 cantidades que se obtiene de la suma de las operaciones de compras y ventas de la hoja de renta fija dolares de gallo. Vemos que rentas y amortizaciones estan ahi pero no se toman en cuenta en el calculo de la cantidad ni del precio costo.   
    * Si vamos al siguiente ejemplo: pampa holding. Representa el otro caso, donde no se tradeo nada en el periodo entonces no podemos calcular el precio costo y aca utilizamos directamente el precio que estaba esa especie al dia de la posicion inicial. Rapidamente nos damos cuenta de que asi como no pudimos obtener los codigos de especie directamente de gallo porque no se tradearon en el periodo de misma forma estos van a tener el precio costo igual al precio de la hoja de posicion final directamente.   
  
Problema 1: cuando no hay movimientos en el periodo del pdf de gallo de alguna especie que aparece en la posicion final no tenemos el codigo de especie entonces el match con los movimientos de visual es por nombre aproximado. Ej: Pampa Pampa Holding en el resumen de gallo vero no sabemos cod especie.   
Se puede solucionar con un listado de especie y codigos general. Tenemos este listado en el archivo “especies” / hoja EspeciesGallo pero se tiene que usar solo como ultimo recurso.   
  
Problema 2: IMPORTANTE: cuando no hay movimientos del activo en el periodo del pdf de gallo, no aparece el movimiento de compra entonces no tenemos el dato del precio del costo inicial. Lo que si tenemos es el precio en la posicion inicial del resumen pdf. Osea 1/1/2025. En estos casos usamos ese precio y agregamos una columna al excel que indica si el precio de extrajo de la hoja de posiciones o si se calculo a traves de la base de datos “origen precio costo”.   
  
## HOJAS AUXILIARES:  
##   
Hoja EspeciesVisual: va a ser fija, es la info para mappear nombre de especie(columna Q) con moneda de emision (col G) y codigo de especie(col A) y Ticker(col H) necesarias para otras hojas   
Hoja EspeciesGallo: va a ser fija, es la info para mappear nombre de especie(columna B) con moneda de emision (col N) y codigo de especie(col A) y ticker(col J) necesarias para otras hojas   
  
Hoja Cotizacion Dolar Historica: Fechas y cotizacion del dolar segun moneda Dolar Mep(local) y Dolar Cable(exterior)   
  
Hoja PreciosInicialesEspecies: listado de las especies de gallo y su valor al 30/12/2024 que se usa como valor inicial en otras hojas.   
##   
## HOJAS DE MOVIMIENTOS:   
Hoja Boletos(OK):   
esta se rellena con la info de las hojas de gallo de movimientos(osea todas menos resultados y posicion inicial/final). En rangos generales la diferencia entre gallo y visual es que gallo divide su estructura de movimientos segun el tipo de instrumento y dentro de un tipo mezcla pesos y dolares. Visual divide su estructura por moneda. La idea es agarrar todos los movimientos de las hojas de gallo y ponerlos en esta hoja.   
Entre parentesis va a que archivo corresponde (v) visual y (g) gallo.   
  
Primero vamos con las columnas que se completan directamente:   
* Col B: Concertacion (v) es la col D: fecha (g) aca hay que filtrar de gallo solo los movimientos de 2025 y tuve que modificar el formato de fecha porque medio que no lo agarraban bien las formulas. A considerar eso   
* COL C:liquidacion(v) en gallo no existe  
* Col D: Nro Boleto(v) es la col F: numero(g)   
* Col F: tipo operacion(v) viene de col E operacion (g). Aca es importante una distincion porque en la hoja de boletos solo se van a poner las operaciones de compra/venta y canje. Las rentas y dividendos van en la hoja rentas dividendos ARS (v) o rentas dividendos USD(v)   
* Col G: Cod. Instrum (v) se completa con los codigos que ya se encontraron para la hoja de posicion final que tmb son los mismos que la col B cod_especie(g)   
* Col H: instrumento crudo (v) es la col C especie(g)  
* Col J: cantidad(v) es la col G cantidad (g)  
* Col K: precio(v) es la col H precio(g)  
  
  
  
Bueno ahora las que tienen tratamiento o que se calculan dsp de haber completado las anteriores. Estan en orden de calculo.   
  
* Col E: Moneda (v) si tiene la columna K:resultado_pesos(g) o la M:gastos_pesos(g) con valor distinto a cero entonces se autocompleta con pesos, si tiene la columna L:resultado_usd(g) o la col N:gastos_usd(g) con valor distinto a cero se completa con Dolar MEP o Dolar Cable. Dolar MEP si el movimiento viene de la hoja tit.privados exentos(g) o de renta fija dolares(g) si la y Dolar Cable si viene de la hoja tit.privados exterior(g).   
  
* Col A: Tipo de Instrumento(v). Si el movimento viene de la hoja tit.privados.exentos y en la columna especie incluye la palabra cedear entonces se completa con cedears ese movimeinto. Si tiene la palabra bono entonces es un titulo publico. Si tiene la palabra ON Es una obligacion negociable. Si tiene la palabra lt o letra o letras es un letras del tesoro nac. Aca se hace un match a traves de la columnma cod.Instrum con la hoja de especieVisual columna codigo y se extrae el valor de la columna Tipo de Especie para colocar aca.   
* Col L: tipo cambio (v) si la moneda es pesos es siempre 1, si es Dolar Mep o Dolar Cable hay que ir a la tabla de precio historico de cada dolar. Uso esta formula =SI(E2="Pesos";1;SI.ERROR(INDICE('Cotizacion Dolar Historica'!$B:$B;COINCIDIR(1;('Cotizacion Dolar Historica'!$A:$A=$B2)*(‘Cotizacion Dolar Historica'!$C:$C=$E2);0));""))  
* Col M: bruto(v) es cantidad(v) x precio(v)  
* Col N:interes(v) siempre cero aparentemente, no tiene par en gallo  
* Col O: gastos(v) se obtiene de la col M gastos_pesos(g) si la moneda es pesos y si es dolar cable o mep de la columna N gastos_usd(g)   
* Col P Neto(v) tiene esta formula siendo I cantidad(v), J precio(v) y N gastos(v).   
    * =SI(I2>0;I2*J2+N2;I2*J2-N2)  
* Col I instrumentoConMoneda(v) aca busco el cod.instrumento en la hoja de EspeciesVisual y haciendo match completo. Esta esta formula =BUSCARV(G2;EspeciesVisual!C:Q;15;FALSO)   
* Agregamos una columna que es Origen: visual o gallo-(nombre hoja)  
* Otra nueva col: moneda emision tiene esta formula =BUSCARV(G2;EspeciesVisual!C:Q;5;FALSO)  
  
HOJA RENTAS Y DIVIDENDOS GALLO(OK)  
Tiene exactamente la misma estructura que boletos solo que el tipo de operacion que se extrae de las hojas de gallo se filtra por Renta, Dividendos y amortizaciones en vez de compras y ventas. Se trae una columna nueva que es costo directamente de gallo homonima. Despues todas las columnas se calculan exactamente igual. Excepto la columna Neto Calculado que es solamente bruto-gastos excepto cuando el tipo de operacion col es amortizacion entonces ahi es -bruto-costo=neto calculado. Despues tenemos la columna origen dice de que hoja de gallo vino.  Se estandariza los valores de moneda, dolar mep o similares a Dolar MEP (local) y dolar cable o similares a Dolar Cable (exterior). Si el col tipo operacion es AMORTIZACION el precio de gallo suele ser 100 pero se transforma a 1 y aca vas a ver que en gallo vienen con costo pero en visual  no y esta bien eso.   
  
## HOJAS DE RESULTADOS:   
En estas 4 hojas basicamente se distribuyen las operaciones de boletos y se les asigna un resultado por cada movimiento.   
  
HOJA RESULTADOS VENTAS ARS (OK)  
* Aca van los movimientos que tienen la palabra compra o venta o cpra o licitaciones pero que la columna InstrumentosConMoneda al final de la palabra dice Pesos. Todas estas columnas se extraen directo a parti de este filtro recien mencionado de la hoja de boletos:   
    * Col: tipo de instrumento   
    * Col: Instrumento  
    * Col: Cod. instrum  
    * Col: Concertacion  
    * Col: liquidacion  
    * Col: moneda   
    * Col: tipo operacion  
    * Col: cantidad  
    * Col: precio  
    * Col: bruto  
    * Col: interes  
    * Col: tipo de cambio   
    * Col: gastos  
* Col: IVA por ahora no la completamos  
* Col: Resultado tampoco  
* Col: cantidad stock (gallo)   
    * Para movimientos que vienen de visual: aca matcheo por cod.instrumento con la hoja de Posicion Final Gallo y agarro valor de la col cantidad   
    * Para movs que vienen de gallo aca deberia obtener la cantidad inicial y precio a partir de la posicion inicial de gallo. hacer: estructurar la hoja de posicion inicial igual que posicion final para que cumpla este objetivo.  
* Col: precio stock(gallo) aca para movimientos que vienen de visual matcheo por cod.instrumento con la hoja de Posicion Final Gallo y agarro el valor de la col ‘precio a utilizar’ pero para movs que vienen de gallo deberia obtener la cantidad inicial y precio a partir de la posicion inicial de gallo. hacer: estructurar la hoja de posicion inicial igual que posicion final para que cumpla este objetivo.  
* Col: costo por venta(gallo) solo para col tipo de operaciones = venta esto es cantidad x precio stock inicial(gallo)  
* Col: Neto Calculado(visual) solo para col tipo de operaciones = venta esto es col bruto+col gastos  
* Col: Resultado Calculado(final) solo para col tipo de operaciones = venta  esto es abs(neto calculado final)- abs(costo por venta)   
* Col: cantidad stock final aca hay que llevar base de datos del stock osea si vendi tengo menos stock y si compre tengo mas stock de esa especie.   
* Col: precio stock final aca hay que llevar base de datos del precio promedio del stock que tengo. Con las ventas queda igual y con las compras cambia.   
* Hay una columna que es la X ‘chequeado’ que tiene comentarios que son importantes para el desarrollo dela logica. Leerlos antes de implementar logica en esta hoja ya que hay casos especiales.   
*   
  
HOJA RESULTADOS VENTAS USD   
* Aca van los movimientos que tienen la palabra compra o venta o cpra o licitaciones pero que la columna InstrumentosConMoneda al final de la palabra dice incluye la palabra Dolar.   
* Aca hay una diferencia en la columna de tipo cambio entre hoja boletos y esta. La primera el tipo de cambio esta con referencia al peso como valor 1 y en la de resultados USD esta con referencia al USD como valor 1. Osea se traduce como loquedigaladeboletos/valor usd de ese dia   
* Mirar bien los comentarios de la columna AA son super importantes.   
  
HOJA RENTAS DIVIDENDOS ARS   
* Aca van los movimientos de la hoja de rentas y dividendos gallo que tienen en la columna moneda de emision : pesos. El match sigue la misma logica que la distribuicion de columnas que se hizo desde boletos hacia resultado ventas ars y usd. Solo que esta tiene solo 4 columnas numericas: cantidad, tipo de cambio, gastos e importe. Este ultimo es el neto calculado de la hoja origen.   
  
HOJA RENTAS DIVIDENDOS USD   
* Misma logica que la hoja rentas dividendos ars pero filtrando por moneda de emision que incluya la palabra dolar en la celdas.   
  
HOJA RESUMEN   
* Es la sumatoria de la columna resultado calculado(final) y de importe dividido entre rentas y amortizaciones como rentas y dividendos por cada hoja de resultados. Viendo la hoja de resumen estan las formulas.   
  
  
Aclaraciones generales: cada vez que haya un cod.instrumento con un punto o ceros adelante hay que sacarselos, siempre son numeros enteros sin punto ni coma. Todas las fechas estructurarlas como fechas.   
  
  
  
  
