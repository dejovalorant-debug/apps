import os, sqlite3, shutil
from datetime import date, datetime
from pathlib import Path
from kivy.app import App
from kivy.core.window import Window
from kivy.metrics import dp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.spinner import Spinner
from kivy.graphics import Color, RoundedRectangle

BG=(.95,.97,.97,1); NAVY=(.07,.22,.26,1); TEAL=(.07,.63,.60,1)
WHITE=(1,1,1,1); MUTED=(.40,.51,.54,1); RED=(.75,.25,.29,1); GREEN=(.12,.52,.42,1)

def money(v):
    s=f"{float(v or 0):,.2f}".replace(",","X").replace(".",",").replace("X",".")
    return "$ "+s

class Card(BoxLayout):
    def __init__(self, **kw):
        super().__init__(orientation="vertical", padding=dp(12), spacing=dp(5), **kw)
        with self.canvas.before:
            Color(*WHITE); self.bg=RoundedRectangle(radius=[dp(12)])
        self.bind(pos=self._sync,size=self._sync)
    def _sync(self,*_): self.bg.pos=self.pos; self.bg.size=self.size

class DeudoresApp(App):
    title="Sistema de Deudores"
    def build(self):
        Window.clearcolor=BG
        self.dbpath=os.path.join(self.user_data_dir,"sistema_deudores.db")
        self.db=sqlite3.connect(self.dbpath); self.db.row_factory=sqlite3.Row
        self.init_db()
        root=BoxLayout(orientation="vertical", spacing=0, padding=dp(4))
        header=BoxLayout(orientation="vertical", size_hint_y=None, height=dp(70),
                         padding=[dp(12),dp(8)], spacing=dp(2))
        with header.canvas.before:
            Color(*NAVY); header_bg=RoundedRectangle(radius=[dp(14),dp(14),0,0])
        header.bind(pos=lambda w,v: setattr(header_bg,"pos",v),
                    size=lambda w,v: setattr(header_bg,"size",v))
        title=Label(text="[b]SISTEMA DE DEUDORES[/b]",markup=True,color=WHITE,
                    font_size="18sp",halign="left",valign="middle",size_hint_y=.55)
        title.bind(size=lambda w,v: setattr(w,"text_size",v))
        subtitle=Label(text="GESTIÓN DE CLIENTES Y CUENTAS",color=(.78,.88,.90,1),
                       font_size="11sp",halign="left",valign="middle",size_hint_y=.45)
        subtitle.bind(size=lambda w,v: setattr(w,"text_size",v))
        header.add_widget(title); header.add_widget(subtitle)
        root.add_widget(header)
        self.body=BoxLayout()
        root.add_widget(self.body)
        nav=GridLayout(cols=4,size_hint_y=None,height=dp(58),spacing=dp(2),padding=[dp(2),dp(3)])
        with nav.canvas.before:
            Color(*NAVY); nav_bg=RoundedRectangle(radius=[0,0,dp(14),dp(14)])
        nav.bind(pos=lambda w,v: setattr(nav_bg,"pos",v),size=lambda w,v: setattr(nav_bg,"size",v))
        self.nav_buttons={}
        for key,icon,label in [("panel","⌂","Panel"),("clientes","♧","Clientes"),
                               ("ventas","▣","Ventas"),("pagos","▤","Pagos")]:
            b=Button(text=f"{icon}\\n{label}",background_normal="",background_color=(0,0,0,0),
                     color=WHITE,font_size="12sp",halign="center",valign="middle")
            b.bind(size=lambda w,v: setattr(w,"text_size",v))
            b.bind(on_release=lambda btn,k=key:self.show(k))
            nav.add_widget(b); self.nav_buttons[key]=b
        root.add_widget(nav)
        self.show("panel")
        return root

    def init_db(self):
        c=self.db.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS clients(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT NOT NULL UNIQUE,phone TEXT DEFAULT '',address TEXT DEFAULT '',notes TEXT DEFAULT '')")
        c.execute("CREATE TABLE IF NOT EXISTS sales(id INTEGER PRIMARY KEY AUTOINCREMENT,sale_date TEXT NOT NULL,receipt TEXT DEFAULT '',client_id INTEGER NOT NULL,detail TEXT DEFAULT '',amount REAL NOT NULL,notes TEXT DEFAULT '')")
        c.execute("CREATE TABLE IF NOT EXISTS payments(id INTEGER PRIMARY KEY AUTOINCREMENT,payment_date TEXT NOT NULL,client_id INTEGER NOT NULL,method TEXT DEFAULT 'Efectivo',amount REAL NOT NULL,notes TEXT DEFAULT '')")
        self.db.commit()
        if not self.db.execute("SELECT COUNT(*) FROM clients").fetchone()[0]:
            x=Path(__file__).parent/"assets"/"datos_iniciales.xlsx"
            if x.exists():
                try: self.import_excel(str(x),quiet=True)
                except Exception: pass

    def import_excel(self,path,quiet=False):
        import openpyxl
        wb=openpyxl.load_workbook(path,data_only=True,read_only=True); clients={}
        if "Clientes" in wb.sheetnames:
            for r in wb["Clientes"].iter_rows(min_row=4,values_only=True):
                n=str(r[1] or "").strip() if len(r)>1 else ""
                if n:
                    self.db.execute("INSERT OR IGNORE INTO clients(name,phone,address,notes) VALUES(?,?,?,?)",(n,str(r[2] or ""),str(r[3] or ""),str(r[4] or "")))
                    clients[n]=self.db.execute("SELECT id FROM clients WHERE name=?",(n,)).fetchone()[0]
        for sheet,kind in [("Ventas","sale"),("Pagos","payment")]:
            if sheet not in wb.sheetnames: continue
            for r in wb[sheet].iter_rows(min_row=4,values_only=True):
                if kind=="sale":
                    if len(r)<5 or not r[2] or not r[4]: continue
                    name=str(r[2]).strip(); amt=float(r[4] or 0); dt=r[0]
                    if name not in clients:
                        self.db.execute("INSERT OR IGNORE INTO clients(name) VALUES(?)",(name,)); clients[name]=self.db.execute("SELECT id FROM clients WHERE name=?",(name,)).fetchone()[0]
                    if amt>0:self.db.execute("INSERT INTO sales(sale_date,receipt,client_id,detail,amount,notes) VALUES(?,?,?,?,?,?)",((dt.strftime("%Y-%m-%d") if isinstance(dt,datetime) else str(dt or date.today())[:10]),str(r[1] or ""),clients[name],str(r[3] or ""),amt,str(r[5] or "")))
                else:
                    if len(r)<4 or not r[1] or not r[3]: continue
                    name=str(r[1]).strip(); amt=float(r[3] or 0); dt=r[0]
                    if name not in clients:
                        self.db.execute("INSERT OR IGNORE INTO clients(name) VALUES(?)",(name,)); clients[name]=self.db.execute("SELECT id FROM clients WHERE name=?",(name,)).fetchone()[0]
                    if amt>0:self.db.execute("INSERT INTO payments(payment_date,client_id,method,amount,notes) VALUES(?,?,?,?,?)",((dt.strftime("%Y-%m-%d") if isinstance(dt,datetime) else str(dt or date.today())[:10]),clients[name],str(r[2] or "Efectivo"),amt,str(r[4] or "")))
        self.db.commit(); wb.close()

    def rows(self):
        return self.db.execute("""SELECT c.*,COALESCE((SELECT SUM(amount) FROM sales s WHERE s.client_id=c.id),0)-COALESCE((SELECT SUM(amount) FROM payments p WHERE p.client_id=c.id),0) balance FROM clients c ORDER BY c.name COLLATE NOCASE""").fetchall()
    def label(self,text,size=14,color=NAVY,bold=False,**kw):
        return Label(text=str(text),font_size=f"{size}sp",color=color,bold=bold,halign="left",valign="middle",**kw)
    def button(self,text,fn):
        b=Button(text=text,size_hint_y=None,height=dp(42),background_normal="",background_color=TEAL,color=WHITE)
        b.bind(on_release=lambda *_:fn()); return b
    def show(self,tab):
        self.body.clear_widgets()
        for k,b in self.nav_buttons.items():
            b.background_color=(.02,.12,.14,1) if k==tab else (0,0,0,0)
        sc=ScrollView(do_scroll_x=False,bar_width=dp(3))
        content=BoxLayout(orientation="vertical",padding=[dp(8),dp(10)],spacing=dp(8),size_hint_y=None)
        content.bind(minimum_height=content.setter("height"))
        sc.add_widget(content); self.body.add_widget(sc)
        if tab=="panel": self.panel(content)
        elif tab=="clientes": self.clients(content)
        elif tab=="ventas": self.records(content,"sales")
        else:self.records(content,"payments")
    def add_metric(self,parent,title,value,color=TEAL):
        card=Card(size_hint_y=None,height=dp(76),padding=dp(11),spacing=dp(3))
        card.add_widget(self.label(title,10,MUTED))
        card.add_widget(self.label(value,19,color,True))
        parent.add_widget(card)
    def panel(self,p):
        rows=self.rows()
        sales=self.db.execute("SELECT COALESCE(SUM(amount),0) FROM sales").fetchone()[0]
        pays=self.db.execute("SELECT COALESCE(SUM(amount),0) FROM payments").fetchone()[0]
        debt=sum(max(0,float(x["balance"])) for x in rows)
        grid=GridLayout(cols=2,spacing=dp(7),size_hint_y=None,height=dp(159))
        for title,val,col in [("VENTAS",money(sales),NAVY),("PAGOS",money(pays),GREEN),
                              ("DEUDA",money(debt),RED),("CLIENTES",len(rows),NAVY)]:
            self.add_metric(grid,title,val,col)
        p.add_widget(grid)
        debtors=sorted([x for x in rows if x["balance"]>0],key=lambda x:-x["balance"])[:5]
        card=Card(size_hint_y=None,height=dp(142),padding=dp(12),spacing=dp(3))
        card.add_widget(self.label("Top 5 deudores",15,NAVY,True))
        for i in range(5):
            if i<len(debtors):
                t=f"{i+1:02d} · {debtors[i]['name']} — {money(debtors[i]['balance'])}"
                col=RED
            else:
                t=f"{i+1:02d} · —"; col=MUTED
            card.add_widget(self.label(t,12,col,size_hint_y=None,height=dp(19)))
        p.add_widget(card)
        owing=sum(x["balance"]>0 for x in rows); ok=len(rows)-owing
        status=Card(size_hint_y=None,height=dp(76))
        status.add_widget(self.label("Estado de cuentas",14,NAVY,True))
        status.add_widget(self.label(f"Con deuda: {owing}     ·     Al día: {ok}",12,GREEN))
        p.add_widget(status)
        p.add_widget(self.button("+ Registrar venta",lambda:self.form("sale")))
        p.add_widget(self.button("+ Registrar pago",lambda:self.form("payment")))

    def clients(self,p):
        p.add_widget(self.label("Clientes",22,NAVY,True))
        p.add_widget(self.button("+ Nuevo cliente",lambda:self.form("client")))
        for c in self.rows():
            card=Card(size_hint_y=None,height=dp(86))
            card.add_widget(self.label(c["name"],15,NAVY,True))
            card.add_widget(self.label(f"{c['phone']}   •   {money(c['balance'])}   •   {'Con deuda' if c['balance']>0 else 'Al día'}",12,RED if c["balance"]>0 else GREEN))
            p.add_widget(card)
    def records(self,p,kind):
        is_sale=kind=="sales"; title="Registro de ventas" if is_sale else "Registro de pagos"
        p.add_widget(self.label(title,22,NAVY,True))
        p.add_widget(self.button("+ Registrar "+("venta" if is_sale else "pago"),lambda:self.form("sale" if is_sale else "payment")))
        q=("SELECT s.sale_date fecha,c.name cliente,s.detail detalle,s.amount importe FROM sales s JOIN clients c ON c.id=s.client_id ORDER BY s.id DESC" if is_sale else "SELECT p.payment_date fecha,c.name cliente,p.method detalle,p.amount importe FROM payments p JOIN clients c ON c.id=p.client_id ORDER BY p.id DESC")
        for r in self.db.execute(q):
            card=Card(size_hint_y=None,height=dp(76))
            card.add_widget(self.label(f"{r['fecha']}  •  {r['cliente']}",14,NAVY,True))
            card.add_widget(self.label(f"{r['detalle']}    {money(r['importe'])}",12,MUTED))
            p.add_widget(card)
    def form(self,kind):
        fields=[]
        layout=BoxLayout(orientation="vertical",padding=dp(14),spacing=dp(8))
        title={"client":"Nuevo cliente","sale":"Registrar venta","payment":"Registrar pago"}[kind]
        layout.add_widget(self.label(title,18,NAVY,True,size_hint_y=None,height=dp(32)))
        def entry(hint):
            x=TextInput(hint_text=hint,multiline=False,size_hint_y=None,height=dp(42),padding=dp(8))
            layout.add_widget(x); fields.append(x); return x
        if kind=="client":
            entry("Nombre y apellido");entry("Teléfono (opcional)");entry("Dirección (opcional)");entry("Notas (opcional)")
        else:
            clients=self.db.execute("SELECT id,name FROM clients ORDER BY name").fetchall()
            if not clients:
                self.popup("Primero registrá un cliente.");return
            sp=Spinner(text=clients[0]["name"],values=[c["name"] for c in clients],size_hint_y=None,height=dp(42))
            layout.add_widget(sp)
            if kind=="sale":entry("Comprobante (opcional)");entry("Detalle / prendas")
            entry("Importe (ej. 15000,50)")
            if kind=="payment":
                method=Spinner(text="Efectivo",values=["Efectivo","Transferencia","Débito","Crédito","Otro"],size_hint_y=None,height=dp(42));layout.add_widget(method)
        row=BoxLayout(size_hint_y=None,height=dp(44),spacing=dp(8))
        pop=Popup(title=title,content=layout,size_hint=(.85,.88))
        def save(*_):
            try:
                if kind=="client":
                    name=fields[0].text.strip()
                    if not name:raise ValueError("Ingresá el nombre.")
                    self.db.execute("INSERT INTO clients(name,phone,address,notes) VALUES(?,?,?,?)",(name,fields[1].text,fields[2].text,fields[3].text))
                else:
                    cid=next(c["id"] for c in clients if c["name"]==sp.text)
                    raw=fields[-1].text.replace("$","").replace(" ","")
                    val=float(raw.replace(".","").replace(",",".") if "," in raw else raw)
                    if val<=0:raise ValueError("El importe debe ser mayor que cero.")
                    if kind=="sale":
                        self.db.execute("INSERT INTO sales(sale_date,receipt,client_id,detail,amount) VALUES(?,?,?,?,?)",(date.today().isoformat(),fields[0].text,cid,fields[1].text,val))
                    else:
                        self.db.execute("INSERT INTO payments(payment_date,client_id,method,amount) VALUES(?,?,?,?)",(date.today().isoformat(),cid,method.text,val))
                self.db.commit();pop.dismiss();self.show("panel" if kind!="client" else "clientes")
            except Exception as e:self.popup(str(e))
        row.add_widget(self.button("Guardar",save)); row.add_widget(self.button("Cancelar",pop.dismiss));layout.add_widget(row)
        pop.open()
    def popup(self,msg):
        box=BoxLayout(orientation="vertical",padding=dp(12));box.add_widget(self.label(msg,14,NAVY))
        pop=Popup(title="Aviso",content=box,size_hint=(.75,.35));box.add_widget(self.button("Aceptar",pop.dismiss));pop.open()
    def on_stop(self):
        try:self.db.close()
        except:pass

if __name__=="__main__": DeudoresApp().run()
