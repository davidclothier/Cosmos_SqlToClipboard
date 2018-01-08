[<img src="http://www.base100.com/images/productos/cosmos.png" alt="Cosmos" width="72"/>](http://www.base100.com/es/productos/cosmos01.html)

# Sql to Clipboard include
Include para copiar al portapapeles de Windows el resultado de una consulta a la base de datos.

## Uso

* La clase ``cSqlToClipboard`` contiene toda la funcionalidad.
* Es necesario **inicializar el objeto** con el método ``reset`` antes de usarlo.
* La única propiedad obligatoria del objeto es la consulta SQL que se asigna con el setter ``setSelect``
* Se puede añadir **una cabecera** al resultado de la consulta con el setter ``setHeader``. No es obligatorio su uso.
* El **separador de campo** por defecto es el carácter *pipeline* ``|``, aunque puede ser cambiado con el setter ``setColumnSeparator``
* El proceso es síncrono. Sin embargo, se puede tener un **callback** para seguir el progreso de la copia mediante el setter ``setHandleCallBack`` que recibe el handle de la ventana que captura el evento ``ExMessage`` de la clase ``cForm``. El identificador del mensaje se obtiene con el método ``getProgressMessageId``. En caso de usar el callback, el objeto instanciado deberá estar declarado en la sección *Variables* u *Objects* para poder acceder al identificador en el evento ``ExMessage`` (o bien copiar ese valor en otra variable una vez inicializado el objeto).
* Es posible cancelar el proceso de copia mediante el método ``cancelQuery`` del objeto. Se copiarán al portapapeles todos los registros hasta el momento de la cancelación.
* El método ``copy`` es el que inicia el cursor sobre la consulta y la copia al portapapeles. Devuelve el número de filas copiadas.

## Limitaciones

* Cada columna de la select tiene que tener como máximo 10.000 caracteres (ya sea una columna individual o la concatenación de varias)

## Ejemplos

### Un ejemplo sencillo
> La base de datos ``stock.dbs`` está copiada en el directorio raíz del proyecto
```
private testSimple
objects
begin
    SqlClipboard    as cSqlToClipboard
    chrSQL          as char default ''
    intRegistros    as integer default 0
end
begin
    PutEnv("DBPATH", ProjectDir());
    Sql.Connect("stock");

    chrSQL = "SELECT item,"
                 +"  supplier,"
                 +"  description,"
                 +"  retail_price,"
                 +"  cost_price,"
                 +"  stock,"
                 +"  last_entry_date"
            +"FROM   ITEMS";

    SqlClipboard.init();
    SqlClipboard.setSelect( chrSQL );
    intRegistros = SqlClipboard.copy();

    MessageBox("Se han copiado "+intRegistros.Using(1)+" registros",Frame.Text,64);
end
```

### Un ejemplo con cabecera, callback y cancelación del proceso
> La base de datos ``stock.dbs`` está copiada en el directorio raíz del proyecto

> El objeto ``mSqlClipboard`` está declarado en la sección *Variables* del módulo
```
private testCabeceraSeparadorYCallback
objects
begin
    chrSQL          as char default ''
    chrSeparator    as char default ''
    chrCabecera     as char default ''
end
begin
    PutEnv("DBPATH", ProjectDir());
    Sql.Connect("stock");

    chrSeparator = 9.Character(); // TAB como separador de campo

    chrSQL = "SELECT item,"
                 +"  supplier,"
                 +"  description,"
                 +"  retail_price,"
                 +"  cost_price,"
                 +"  stock,"
                 +"  last_entry_date"
            +"FROM   ITEMS";

    chrCabecera = "Item"+chrSeparator
                 +"Supplier"+chrSeparator
                 +"Description"+chrSeparator
                 +"Retail Price"+chrSeparator
                 +"Cost Price"+chrSeparator
                 +"Stock"+chrSeparator
                 +"Last Entry Date";

    mSqlClipboard.init();
    mSqlClipboard.setColumnSeparator( chrSeparator );
    mSqlClipboard.setSelect( chrSQL );
    mSqlClipboard.setHeader( chrCabecera );
    mSqlClipboard.setHandleCallBack( self.Hwnd() ); // Recibimos en la ventana actual el progreso de la copia al portapapeles
    mSqlClipboard.copy();
end

On ExMessage (wParam as integer,lParam as integer)
begin
    // Capturamos la información de progreso de la exportación. lParam recibe el número de fila
    if wParam == mSqlClipboard.getProgressMessageId() then
    begin
        txtInfo.Text = lParam;
        Yield();

        // Cancelamos el proceso en la fila 150
        if lParam == 150 then mSqlClipboard.cancelQuery();
    end
end
```

## Mejoras futuras

* Obtener de las *systables/syscolumns* las labels o nombres de columnas para generar la cabecera automáticamente si no se indica

## Contacto

* David Ropero Alcázar - david.ropero@gmail.com

[<img src="https://image.freepik.com/iconos-gratis/boton-del-logotipo-de-twitter_318-85053.jpg" alt="Twitter" width="32"/>](https://twitter.com/clothierdroid)
[<img src="http://www.iconsdb.com/icons/preview/black/linkedin-4-xxl.png" alt="Linkedin" width="32"/>](http://www.linkedin.com/in/davidropero)

> Porque ya era hora de que alguien subiera un repo en Cosmos a github :stuck_out_tongue_winking_eye:

## Licencia

GNU Public License v3.0
