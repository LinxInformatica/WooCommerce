# coding: UTF-8


from woocommerce import API
import pyodbc
import datetime
import argparse
import datetime
import tkinter as tk
from tkinter import ttk
from tkinter import Listbox
import tkinter.scrolledtext as st
from  tkinter import *
from  urllib import request
import threading

class Aplicacion():

    def __init__(self):
        super().__init__()
        self.ventana = tk.Tk()
        self.ventana.title("Api de Comercio Electrónico")
        self.ventana.geometry("500x500")
        self.ventana.iconbitmap("linx.ico")

        self.listaprogreso = tk.Listbox(self.ventana)
        self.listaprogreso.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        scrollbar = Scrollbar(self.listaprogreso)

        self.boton =tk.Button(self.ventana, text="Cerrar", command=self.salir, state=DISABLED)
        scrollbar.pack(side=RIGHT, fill=BOTH)
        self.boton.pack(fill=tk.X, padx=10, pady=10)
        self.listaprogreso.config(yscrollcommand = scrollbar.set)
        scrollbar.config(command = self.listaprogreso.yview)

    def salir(self):
        self.ventana.destroy()

class eCommerce(object):
    def __init__(self):
        super().__init__()
        self.procesocompleto=True
        self.accion = ""
        self.idGrupo = ""
        self.cadenaodbc = ""
        self.idApi = ""
        self.idArticulo=""
        self.idGrupo=""
        self.idOrden=""
        self.idCliente=""
        self.paginado=0
        self.timeout=0
        self.cursor = pyodbc.Cursor
        self.wcapi = API
        self.app = Aplicacion()
        self.iniciar()

    def controlhilo(self,hilo):
        self.app.ventana.after(1000,self.controlfinhilo,hilo)

    def controlfinhilo(self,hilo):
        if hilo.is_alive():
            self.controlhilo(hilo)

    def salir(self):
        self.app.ventana.destroy()

    def iniciar(self):
        self.funcionhilo()

        self.app.ventana.mainloop()

    def funcionhilo(self):
        hilo = threading.Thread(target=self.opciones())
        hilo.start()
        self.controlhilo(hilo)

    def opciones(self):
        parser = argparse.ArgumentParser()
        parser.add_argument("-o", "--odbc", help="Cadena odbc",default="DSN=margarita")
        parser.add_argument("-a", "--api", help="id Api",default="1")
        parser.add_argument("-i", "--importar", help="Importar",default=None)
        parser.add_argument("-e", "--exportar", help="Exportar",default=None)
        parser.add_argument("-l", "--listar", help="Listar",default="1")
        args = parser.parse_args()
        self.cadenaConeccion=args.odbc
        self.idApi=args.api
        if args.importar != None:
            if self.ConeccionODBC() and self.ConeccionSitio():
                if args.importar[0:1]=="A" or args.importar[0:1]=="F":
                    #ImportarArticulos"
                    self.idArticulo=args.importar[1:]
                    if self.idArticulo=="":
                        self.Articulos_get()
                    else:
                        self.Articulo_get()

                if args.importar[0:1]=="G" or args.importar[0:1]=="F":
                    #ImportarGrupos
                    self.idGrupo=args.importar[1:]

                if args.importar[0:1]=="O" or args.importar[0:1]=="F":
                    #ImportarOrdenes
                    self.Ordenes_get()


        elif args.exportar != None:
            if self.ConeccionODBC() and self.ConeccionSitio():
                if args.exportar[0:1]=="A" or args.exportar[0:1]=="F":
                    #ExportarArticulos
                    self.idArticulo=args.exportar[1:]
                    self.Articulo_put()

                if args.exportar[0:1]=="G" or args.exportar[0:1]=="F":
                    #ExportarGrupos
                    self.idGrupo=args.exportar[1:]
                    self.Grupo_put()

                if args.exportar[0:1]=="O" or args.exportar[0:1]=="F":
                    #ExportarOrdenes
                    self.idOrden=args.exportar[1:]

                if args.exportar[0:1]=="C" or args.exportar[0:1]=="F":
                    #ExportarClientes
                    self.idCliente=args.exportar[1:]
                    self.Cliente_put()
        elif  args.listar != None:
            if self.ConeccionODBC() and self.ConeccionSitio():
                self.Articulos_Listar()
        else:
            self.error("Parametros incorrectos")

        if self.procesocompleto:
            self.progreso("Proceso Finalizado OK")
        self.app.boton.configure(state=NORMAL)

    def progreso(self,mensaje):
        self.app.listaprogreso.insert(END,mensaje)
        self.app.listaprogreso.select_clear(self.app.listaprogreso.size() - 2)   #Clear the current selected item
        self.app.listaprogreso.select_set(END)                             #Select the new item
        self.app.listaprogreso.yview(END)
        self.app.ventana.update()

    def error(self,mensaje="Error"):
        self.procesocompleto=False
        self.progreso(mensaje)
        self.progreso("Proceso finalizado con Error!!!!")

    def ConeccionODBC(self):
        try:
            self.progreso("Conectando a "+self.cadenaConeccion)
            cnxn = pyodbc.connect(self.cadenaConeccion, autocommit=True)
            self.cursor = cnxn.cursor()
            return True
        except:
            self.error('No existe la cadena de coneccion:'+self.cadenaConeccion)
            return False

    def ConeccionSitio(self):

        cadenaSQL= """
                   select  sitio, 
                           clavecliente,
                           clavesecreta,
                           {fn convert(timeout,SQL_INTEGER)} as timeout,
                           {fn convert(paginado,SQL_INTEGER)} as paginado
                           from apiCabezera 
                           where idinternoapicabezera= ?
                   """

        self.cursor.execute(cadenaSQL,self.idApi)

        row = self.cursor.fetchone()
        if row:
            self.progreso("Accediendo a  "+row.sitio)
            try:
                request.urlopen(row.sitio,timeout=115)
            except:
                self.error("El sitio:"+row.sitio+" es inaccesible, Controle su coneccion a internet")
                return False
            try:
                self.progreso("Chequeando credenciales ")
                self.timeout=row.timeout
                self.paginado=row.paginado
                if self.timeout==0:
                    self.timeout=100
                if self.paginado==0:
                    self.paginado=100

                self.wcapi = API(
                    url=row.sitio,
                    consumer_key=row.clavecliente,
                    consumer_secret=row.clavesecreta,
                    wp_api=True,
                    version="wc/v3",
                    query_string_auth=True,
                    timeout=self.timeout
                    )
            except:
                self.error("No se pudo validar el acceso al sitio "+row.sitio)
                return False
        else:
            self.error("Verifique la configuracin de la api")
            return False
        return True

    def Articulo_put(self):
        cadenaSQL="""select 
                     apidatos.idinternodato,
                     apidatos.descripcion,
                     apidatos.preciopublicado,
                     apidatos.idpublicado,
                     apidatos.sku,
                     {fn convert(apidatos.cantidad,SQL_REAL)} as cantidad,
                     {fn convert(apidatos.activo,SQL_INTEGER)} as activo 
                     from apidatos
                     where idinternoapicabezera = ? 
                     and tipodato=1
                     """
        if self.idArticulo=="":
            self.cursor.execute(cadenaSQL,self.idApi)
        else:
            cadenaSQL=cadenaSQL+" and apidatos.sku=?"
            self.cursor.execute(cadenaSQL,self.idApi,self.idArticulo)
        rows = self.cursor.fetchall()
        listaCreate = []
        listaUpdate = []
        fila=0
        registro=0
        if rows:
            for row in rows:
                self.progreso("Preparando "+row.descripcion)
                if row.activo==1:
                    catalog_visibility='visible'
                else:
                    catalog_visibility='hidden'
                dictItems = {
                            "name": row.descripcion,
                            "short_description": row.descripcion,
                            "regular_price": str(float(row.preciopublicado)),
                            "manage_stock": "true",
                            "stock_quantity": str(int(row.cantidad)),
                            "sku": row.sku,
                            "catalog_visibility": catalog_visibility
                        }
                print(dictItems)
                if row.idpublicado:
                    dictItems["id"] = row.idpublicado

                    listaUpdate.append(dictItems)
                else:
                    #dictItems["id"] = row.sku

                    listaCreate.append(dictItems)
                fila+=1
                registro+=1
                if fila==self.paginado or registro==len(rows):
                    self.Articulos_ActualizoSitio(listaCreate,listaUpdate)
                    listaCreate = []
                    listaUpdate = []
                    fila=0
        else:
            self.error("NO hay articulos para procesar")

    def Articulos_ActualizoSitio(self,listaCreate,listaUpdate):
        data = {}
        if listaCreate:
            data["create"]=listaCreate
        if listaUpdate:
            data["update"]=listaUpdate
        print(data)
        self.progreso("Cargando Articulos al sitio")
        respuesta = self.wcapi.post("products/batch", data).json()
        if listaCreate:
            print(respuesta["create"])
            for item in respuesta['create']:
                self.progreso('Actualizando '+item["name"])
                print(item["name"])
                cadenaSQL="""
                         update apidatos set 
                         idpublicado= ?
                         where sku = ? and
                         idinternoapicabezera = ?
                          """
                parametros=(item['id'],item['sku'],self.idApi)
                self.cursor.execute(cadenaSQL,parametros)

    def Articulos_get(self):
        self.progreso("Importando Articulos...")

        cadenaSQL="""select count(*) as cantidadarticulos
                     from apidatos
                     where
                     idinternoapicabezera = ? 
                     and idpublicado<>''
                     and tipodato=1"""
        self.cursor.execute(cadenaSQL,self.idApi)
        apidatos=self.cursor.fetchone()
        if apidatos.cantidadarticulos ==0:
            self.error("No hay articulos para importar")
            return False
        ciclos=apidatos.cantidadarticulos // self.paginado
        for pag in range(0,ciclos):
            cadenaApi="products"
            try:
                respuesta=self.wcapi.get(cadenaApi,params={"page": pag+1,"per_page": self.paginado}).json()
                for fila in respuesta:
                    self.Articulo_ActualizoSistema(fila)
            except:
                self.error("Error para importar articulos")
                return False
        ciclofinal=int(((apidatos.cantidadarticulos / self.paginado)-(ciclos))*100)
        if ciclofinal:
            cadenaApi="products"
            #try:
            if ciclofinal:
                respuesta=self.wcapi.get(cadenaApi,params={"page": ciclos+1,"self.paginado": ciclofinal}).json()
                print(respuesta)
                for fila in respuesta:
                    self.Articulo_ActualizoSistema(fila)
            #except:
            #    self.error("Error en ciclo final de importacion de articulos")
            #    return False
        return True

    def Articulo_get(self):
        cadenaApi="products/"+self.idArticulo
        self.Articulo_ActualizoSistema(self.wcapi.get(cadenaApi).json())

    def Articulos_Listar(self):
        print(self.wcapi.get("").json())


        cadenaApi="products"
        salida=self.wcapi.get(cadenaApi).json()
        print(salida)

    def Articulo_ActualizoSistema(self,respuesta):
        if 'id' in respuesta:
            if respuesta['catalog_visibility']=='visible':
                activo=1
            else:
                activo=0

            self.progreso("Actualizando "+respuesta['name'])
            cadenaSQL="""
                      update apidatos set 
                         descripcion= ?,
                         preciopublicado = ?,
                         cantidad =?,
                         activo =?
                         where
                         idpublicado = ? and
                         idinternoapicabezera = ?
                              """
            parametros=(respuesta['name'],respuesta['regular_price'],respuesta['stock_quantity'],activo,respuesta['id'],self.idApi)
            self.cursor.execute(cadenaSQL,parametros)

    def Grupo_put(self):
        cadenaSQL="""select idinternodato,
                     apidatos.descripcion,
                     apidatos.idpublicado
                     from apidatos
                     where idinternoapicabezera = ? 
                     and tipodato=2"""
        if self.idGrupo=="":
            self.cursor.execute(cadenaSQL,self.idApi)
        else:
            cadenaSQL=cadenaSQL+" and apidatos.idinternodato=?"
            self.cursor.execute(cadenaSQL,self.idApi,self.idGrupo)

        rows = self.cursor.fetchall()

        data = {}
        listaCreate = []
        listaUpdate = []
        fila=0
        for row in rows:
            fila+=1
            if fila==100:
                break
            dictItems = {
                        "name": row.descripcion,
                        "slug": row.idinternodato
                    }
            if row.idpublicado:
                dictItems["id"] = row.idpublicado

                listaUpdate.append(dictItems)
            else:
                listaCreate.append(dictItems)

        if listaCreate:
            data["create"]=listaCreate
        if listaUpdate:
            data["update"]=listaUpdate

        respuesta = self.wcapi.post("products/categories/batch", data).json()

        if listaCreate:
            for item in respuesta['create']:

                cadenaSQL="""
                         update apidatos set 
                         idpublicado= ?
                         where 
                         idinternodato = ? and
                         idinternoapicabezera = ?                                
                          """
                parametros = (item['id'],item['slug'],self.idApi)
                self.cursor.execute(cadenaSQL,parametros)

    def Cliente_put(self):
        cadenaSQL="""select idinternodato,
                     apidatos.descripcion,
                     apidatos.idpublicado,
                     vclientes.documento,
                     vclientes.email,
                     vclientes.domicilio,
                     vclientes.telefono,
                     glocalidades.descripcion as localidad,
                     glocalidades.codigopostal as codigopostal,
                     gprovincias.nombreprovincia as provincia
                     from apidatos,vclientes,glocalidades,gprovincias
                     where 
                     idinternoapicabezera = ? 
                     and tipodato=4
                     and idinternodato={fn convert(vclientes.idcliente,SQL_CHAR)}
                     and glocalidades.idlocalidad=vclientes.idlocalidad
                     and gprovincias.idprovincia=glocalidades.idprovincia
                      """
        if self.idCliente=="":
            self.cursor.execute(cadenaSQL,self.idApi)
        else:
            cadenaSQL=cadenaSQL+" and apidatos.idinternodato=?"
            self.cursor.execute(cadenaSQL,self.idApi,self.idCliente)

        rows = self.cursor.fetchall()
        if rows:
            data = {}
            listaCreate = []
            listaUpdate = []
            fila=0
            for row in rows:
                fila+=1
                if fila==100:
                    break
                if row.email and "@" in row.email:
                    dictBilling = {
                                "first_name": row.descripcion,
                                "company": row.descripcion,
                                "address_1": row.domicilio,
                                "city": row.localidad,
                                "state": row.provincia,
                                "postcode": row.codigopostal,
                                "phone": row.telefono,
                                "email": row.email
                                }
                    dictItems = {
                                "first_name": row.descripcion,
                                "username": row.idinternodato,
                                "password": row.documento,
                                "email":row.email,
                                "billing":dictBilling,
                                "shipping":dictBilling
                            }
                    if row.idpublicado:
                        dictItems["id"] = int(row.idpublicado)

                        listaUpdate.append(dictItems)
                    else:
                        listaCreate.append(dictItems)
                else:
                    self.error("El cliente "+row.descripcion+" tiene una direccion de mail incorrecta!")
                    return False

            if listaCreate:
                data["create"]=listaCreate
            if listaUpdate:
                data["update"]=listaUpdate

            respuesta = self.wcapi.post("customers/batch", data).json()

            if listaCreate:
                for item in respuesta['create']:
                    if item["id"] != 0:
                        cadenaSQL="""
                                 update apidatos set 
                                 idpublicado= ?
                                 where 
                                 idinternodato = ? and
                                 idinternoapicabezera = ?                                
                                  """
                        parametros = (item['id'],item['username'],self.idApi)
                        self.cursor.execute(cadenaSQL,parametros)
                    else:
                        erroritem=item["error"]
                        self.progreso(erroritem["message"])

        else:
            self.error("No hay Clientes para procesar")

    def Ordenes_get(self):
        cadenaSQL= """
                   select  estadopendiente,
                           idtipocomprobante,
                           iddeposito
                           from apiCabezera 
                           where idinternoapicabezera= ?
                   """
        self.cursor.execute(cadenaSQL,self.idApi)
        apicabezera = self.cursor.fetchone()
        self.progreso("Descargando Ordenes")
        respuesta= self.wcapi.get("orders",params={"status": apicabezera.estadopendiente}).json()
        if respuesta:
            #Busco Proximo comprobante
            cadenaSQL= """
                   select  max(idinterno) as maxidinterno
                           from vcomprobantes 
                   """
            self.cursor.execute(cadenaSQL)
            vcomprobante = self.cursor.fetchone()
            #Busco tipo Comprobante
            cadenaSQL= """
                   select  vtipocomprobantes.idtipocomprobante,
                           vtalonarios.letraasociada,
                           vtalonarios.sucursalasociada,
                           vtipocomprobantes.comportamiento,
                           vtalonarios.origen,
                           vtipocomprobantes.signo
                           from vtipocomprobantes 
                             inner join vtalonarios on vtipocomprobantes.idtalonario=vtalonarios.idtalonario
                           where 
                             vtipocomprobantes.idtipocomprobante=?
                           
                   """
            self.cursor.execute(cadenaSQL,apicabezera.idtipocomprobante)
            vtipocomprobantes = self.cursor.fetchone()

            idinterno=vcomprobante.maxidinterno
            for orden in respuesta:
                #Busco cliente
                cadenaSQL= """
                       select   idcliente,
                                nombrecliente,
                                idcondicionfiscal,
                                tipodocumento,
                                documento,
                                domicilio,
                                idlocalidad,
                                idvendedor,
                                idzona,
                                idtransporte
                                from vclientes 
                            """
                if orden["customer_id"]==0:
                    cadenaSQL=cadenaSQL+"""
                                 where idcliente=?"""
                    parametros=(1)
                else:
                    cadenaSQL=cadenaSQL+"""
                          inner join apidatos on idinternodato={fn convert(vclientes.idcliente,SQL_CHAR)}
                                where idinternoapicabezera = ? 
                                and tipodato=4
                                and apidatos.idpublicado=? 
                       """
                    parametros=(self.idApi,orden["customer_id"])
                self.cursor.execute(cadenaSQL,parametros)
                cliente = self.cursor.fetchone()

                idinterno+=1
                cadenaSQL= """
                   insert into vcomprobantes(
                               idinterno,
                               idcliente,
                               nombrecliente,
                               idcondicionfiscal,
                               tipodocumento,
                               documento,
                               domicilio,
                               idlocalidad,
                               idvendedor,
                               idzona,
                               origen,
                               idtipocomprobante,
                               comprobante_letra,
                               comprobante_suc,
                               comprobante_numero,
                               comportamiento,
                               signo,
                               datoentidadexterna,
                               importetotal,
                               estado,
                               fecha,
                               fechafiscal,
                               fechaa,
                               horaa
                               )  values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) 
                   """
                fechaorden=datetime.datetime.strptime(orden["date_created"],"%Y-%m-%dT%H:%M:%S").date()
                horaorden=datetime.datetime.strptime(orden["date_created"],"%Y-%m-%dT%H:%M:%S").time()
                nombrecliente=""
                domiciliocliente=""
                if "billing" in orden:
                    clientesitio=orden["billing"]
                    nombrecliente=clientesitio["last_name"].upper()+" "+clientesitio["first_name"].upper()
                    domiciliocliente=cliente.domicilio
                if nombrecliente.strip(" ")=="":
                    nombrecliente=cliente.nombrecliente
                if domiciliocliente.strip(" ")=="":
                    domiciliocliente=cliente.domicilio
                parametros=(idinterno,
                            cliente.idcliente,
                            nombrecliente,
                            cliente.idcondicionfiscal,
                            cliente.tipodocumento,
                            cliente.documento,
                            domiciliocliente,
                            cliente.idlocalidad,
                            cliente.idvendedor,
                            cliente.idzona,
                            vtipocomprobantes.origen,
                            vtipocomprobantes.idtipocomprobante,
                            vtipocomprobantes.letraasociada,
                            vtipocomprobantes.sucursalasociada,
                            orden["id"],
                            vtipocomprobantes.comportamiento,
                            vtipocomprobantes.signo,
                            orden["id"],
                            orden["total"],
                            7,
                            datetime.date.today(),
                            datetime.date.today(),
                            fechaorden,
                            horaorden
                            )
                self.progreso("Grabando Orden "+str(orden["id"])+" de "+nombrecliente)
                try:
                    self.cursor.execute(cadenaSQL,parametros) #insert vcomprobantes
                except:
                    self.error("Error de actualizacion")
                    return False

                #
                if "line_items" in orden:
                    #Busco Proximo vitem
                    cadenaSQL= """
                            select  max(idinternoitems) as maxidinterno
                                    from vitems 
                            """
                    self.cursor.execute(cadenaSQL)
                    vitem = self.cursor.fetchone()
                    idinternoitems=vitem.maxidinterno
                    idrenglon=0
                    for items in orden["line_items"]:
                        #Busco articulo
                        cadenaSQL= """
                               select   idinternoarticulo,
                                        idarticulo,
                                        descripcion,
                                        idalicuotaiva,
                                        idcuenta,
                                        idimpuestointerno,
                                        idunidadmedidaventas,
                                        idcentro
                                        from sarticulos 
                                        where idinternoarticulo=? 
                               """
                        parametros=(items["sku"])
                        self.cursor.execute(cadenaSQL,parametros)
                        sarticulos = self.cursor.fetchone()
                        idinternoitems+=1
                        idrenglon+=1
                        cadenaSQL= """
                               insert into vitems(
                                           idinternoitems,
                                           idinterno,
                                           idrenglon,
                                           idinternoarticulo,
                                           idarticulo,
                                           descripcion,
                                           cantidad,
                                           cantidadstock,
                                           preciounitario,
                                           preciounitarioajustado,
                                           preciofactura,
                                           precioneto,
                                           preciofinal,
                                           idalicuotaiva,
                                           idcuenta,
                                           idimpuestointerno,
                                           iddeposito,
                                           idunidadmedida,
                                           idcentro
                                           )  values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) 
                               """
                        parametros=(idinternoitems,
                                    idinterno,
                                    idrenglon,
                                    sarticulos.idinternoarticulo,
                                    sarticulos.idarticulo,
                                    sarticulos.descripcion,
                                    items["quantity"],
                                    items["quantity"],
                                    items["price"],
                                    items["price"],
                                    items["quantity"]*items["price"],
                                    items["quantity"]*items["price"],
                                    items["quantity"]*items["price"],
                                    sarticulos.idalicuotaiva,
                                    sarticulos.idcuenta,
                                    sarticulos.idimpuestointerno,
                                    apicabezera.iddeposito,
                                    sarticulos.idunidadmedidaventas,
                                    sarticulos.idcentro
                                    )
                        try:
                            self.cursor.execute(cadenaSQL,parametros) #insert vitems
                        except:
                            self.error("Error de actualizacion de datos")
def main():
    objeto = eCommerce()
    return 0

if __name__ == "__main__":
    main()

