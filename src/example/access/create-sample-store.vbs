' ============================================================
'  create-sample-store.vbs
'  Creates sample-store.mdb (Microsoft Access) next to this script,
'  with 4 related tables filled with example data + 2 saved queries.
'  Tables, columns, constraints and queries are snake_case, keywords
'  upper case (see CLAUDE.md, SQL convention).
'
'  Usage: double-click this file on Windows.
'  If you get "Provider cannot be found", run it with the 32-bit host:
'     C:\Windows\SysWOW64\wscript.exe create-sample-store.vbs
' ============================================================
Option Explicit

Dim fso, folder, dbPath, connStr, cat, cn, providers, p, ok
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
dbPath = folder & "\sample-store.mdb"

If fso.FileExists(dbPath) Then
    If MsgBox("sample-store.mdb already exists. Replace it?", vbYesNo + vbQuestion, "Sample Store") = vbNo Then WScript.Quit
    fso.DeleteFile dbPath
End If

' Try the Access engine (ACE) first, then the Jet engine built into Windows
providers = Array("Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0", "Microsoft.Jet.OLEDB.4.0")
ok = False
On Error Resume Next
For Each p In providers
    Err.Clear
    connStr = "Provider=" & p & ";Data Source=" & dbPath & ";Jet OLEDB:Engine Type=5;"
    Set cat = CreateObject("ADOX.Catalog")
    cat.Create connStr
    If Err.Number = 0 Then ok = True: Exit For
Next
On Error GoTo 0
If Not ok Then
    MsgBox "No Access/Jet database engine found." & vbCrLf & vbCrLf & _
           "Try running:  C:\Windows\SysWOW64\wscript.exe """ & WScript.ScriptFullName & """" & vbCrLf & _
           "or install the free Microsoft Access Database Engine.", vbExclamation, "Sample Store"
    WScript.Quit 1
End If

Set cn = cat.ActiveConnection

' ---------- Tables ----------
cn.Execute "CREATE TABLE customers (" & _
    "customer_id COUNTER CONSTRAINT pk_customers PRIMARY KEY, " & _
    "first_name TEXT(50) NOT NULL, last_name TEXT(50) NOT NULL, " & _
    "email TEXT(100), phone TEXT(20), city TEXT(50), country TEXT(50), " & _
    "signup_date DATETIME)"

cn.Execute "CREATE TABLE products (" & _
    "product_id COUNTER CONSTRAINT pk_products PRIMARY KEY, " & _
    "product_name TEXT(100) NOT NULL, category TEXT(50), " & _
    "unit_price CURRENCY, units_in_stock INTEGER, discontinued YESNO)"

cn.Execute "CREATE TABLE orders (" & _
    "order_id COUNTER CONSTRAINT pk_orders PRIMARY KEY, " & _
    "customer_id INTEGER NOT NULL, order_date DATETIME, shipped_date DATETIME, " & _
    "status TEXT(20), " & _
    "CONSTRAINT fk_orders_customers FOREIGN KEY (customer_id) REFERENCES customers (customer_id))"

cn.Execute "CREATE TABLE order_details (" & _
    "order_detail_id COUNTER CONSTRAINT pk_order_details PRIMARY KEY, " & _
    "order_id INTEGER NOT NULL, product_id INTEGER NOT NULL, " & _
    "quantity INTEGER, unit_price CURRENCY, discount DOUBLE, " & _
    "CONSTRAINT fk_details_orders FOREIGN KEY (order_id) REFERENCES orders (order_id), " & _
    "CONSTRAINT fk_details_products FOREIGN KEY (product_id) REFERENCES products (product_id))"

' ---------- Data ----------
Dim r, c, pr, o, d, shipped
' first_name|last_name|email|phone|city|country|signup_date(MM/DD/YYYY)
For Each r In Array( _
    "Claire|Martin|claire.martin@example.com|+33 4 50 12 34 56|Annecy|France|01/15/2025", _
    "Lucas|Bernard|lucas.bernard@example.com|+33 1 42 98 76 54|Paris|France|02/03/2025", _
    "Sofia|Rossi|sofia.rossi@example.com|+39 02 1234 5678|Milan|Italy|02/20/2025", _
    "Jonas|Weber|jonas.weber@example.com|+49 30 9876 5432|Berlin|Germany|03/11/2025", _
    "Emma|Dubois|emma.dubois@example.com|+33 4 78 11 22 33|Lyon|France|04/02/2025", _
    "Noah|Keller|noah.keller@example.com|+41 22 555 66 77|Geneva|Switzerland|04/28/2025", _
    "Olivia|Smith|olivia.smith@example.com|+44 20 7946 0011|London|United Kingdom|05/19/2025", _
    "Hugo|Garcia|hugo.garcia@example.com|+34 91 123 45 67|Madrid|Spain|06/07/2025", _
    "Lea|Fontaine|lea.fontaine@example.com|+33 4 91 44 55 66|Marseille|France|07/14/2025", _
    "Mateo|Silva|mateo.silva@example.com|+351 21 987 6543|Lisbon|Portugal|08/30/2025")
    c = Split(r, "|")
    cn.Execute "INSERT INTO customers (first_name, last_name, email, phone, city, country, signup_date) VALUES ('" & _
        c(0) & "','" & c(1) & "','" & c(2) & "','" & c(3) & "','" & c(4) & "','" & c(5) & "',#" & c(6) & "#)"
Next

' product_name|category|unit_price|units_in_stock|discontinued
For Each r In Array( _
    "Wireless Mouse|Accessories|24.90|120|0", _
    "Mechanical Keyboard|Accessories|89.00|45|0", _
    "27"" 4K Monitor|Displays|349.00|18|0", _
    "USB-C Dock|Accessories|129.50|30|0", _
    "Laptop Stand|Furniture|39.90|75|0", _
    "Noise-Cancelling Headset|Audio|199.00|25|0", _
    "Webcam HD 1080p|Video|59.00|60|0", _
    "Ergonomic Chair|Furniture|279.00|12|0", _
    "External SSD 1TB|Storage|109.00|40|0", _
    "VGA Adapter|Accessories|9.90|0|-1")
    pr = Split(r, "|")
    cn.Execute "INSERT INTO products (product_name, category, unit_price, units_in_stock, discontinued) VALUES ('" & _
        pr(0) & "','" & pr(1) & "'," & pr(2) & "," & pr(3) & "," & pr(4) & ")"
Next

' customer_id|order_date|shipped_date|status
For Each r In Array( _
    "1|09/02/2025|09/04/2025|Delivered", "2|09/05/2025|09/08/2025|Delivered", _
    "3|09/12/2025|09/15/2025|Delivered", "1|10/01/2025|10/03/2025|Delivered", _
    "4|10/09/2025|10/13/2025|Delivered", "5|10/21/2025|10/23/2025|Delivered", _
    "6|11/04/2025|11/07/2025|Delivered", "7|11/18/2025|11/20/2025|Delivered", _
    "2|12/01/2025|12/03/2025|Delivered", "8|12/15/2025|12/19/2025|Delivered", _
    "9|01/10/2026|01/13/2026|Delivered", "10|02/02/2026|02/05/2026|Delivered", _
    "3|03/14/2026|03/17/2026|Delivered", "5|09/28/2026|10/01/2026|Shipped", _
    "7|10/10/2026||Pending")
    o = Split(r, "|")
    If o(2) = "" Then shipped = "NULL" Else shipped = "#" & o(2) & "#"
    cn.Execute "INSERT INTO orders (customer_id, order_date, shipped_date, status) VALUES (" & _
        o(0) & ",#" & o(1) & "#," & shipped & ",'" & o(3) & "')"
Next

' order_id|product_id|quantity|unit_price|discount
For Each r In Array( _
    "1|1|2|24.90|0", "1|2|1|89.00|0", "2|3|2|349.00|0.05", "2|4|1|129.50|0", _
    "3|6|1|199.00|0", "3|7|1|59.00|0", "4|9|2|109.00|0.1", "5|8|1|279.00|0", _
    "5|5|1|39.90|0", "6|1|3|24.90|0", "6|5|2|39.90|0", "7|3|1|349.00|0", _
    "7|4|1|129.50|0", "7|2|1|89.00|0", "8|6|2|199.00|0.1", "9|7|4|59.00|0.05", _
    "10|8|2|279.00|0", "10|5|2|39.90|0", "11|9|1|109.00|0", "11|1|1|24.90|0", _
    "12|2|1|89.00|0", "12|6|1|199.00|0", "13|3|3|349.00|0.15", "13|4|3|129.50|0.1", _
    "14|7|1|59.00|0", "14|9|2|109.00|0", "15|1|5|24.90|0", "15|5|1|39.90|0")
    d = Split(r, "|")
    cn.Execute "INSERT INTO order_details (order_id, product_id, quantity, unit_price, discount) VALUES (" & _
        d(0) & "," & d(1) & "," & d(2) & "," & d(3) & "," & d(4) & ")"
Next

' ---------- Saved queries ----------
Dim cmd
Set cmd = CreateObject("ADODB.Command")
cmd.CommandText = "SELECT o.order_id, o.order_date, c.first_name & ' ' & c.last_name AS customer, o.status, " & _
    "SUM(d.quantity * d.unit_price * (1 - d.discount)) AS order_total " & _
    "FROM (orders AS o INNER JOIN customers AS c ON o.customer_id = c.customer_id) " & _
    "INNER JOIN order_details AS d ON o.order_id = d.order_id " & _
    "GROUP BY o.order_id, o.order_date, c.first_name, c.last_name, o.status"
cat.Views.Append "qry_order_totals", cmd

Set cmd = CreateObject("ADODB.Command")
cmd.CommandText = "SELECT p.product_name, p.category, SUM(d.quantity) AS units_sold, " & _
    "SUM(d.quantity * d.unit_price * (1 - d.discount)) AS revenue " & _
    "FROM products AS p INNER JOIN order_details AS d ON p.product_id = d.product_id " & _
    "GROUP BY p.product_name, p.category"
cat.Views.Append "qry_sales_by_product", cmd

cn.Close
MsgBox "Done!" & vbCrLf & dbPath & vbCrLf & vbCrLf & _
       "Tables: customers (10), products (10), orders (15), order_details (28)" & vbCrLf & _
       "Queries: qry_order_totals, qry_sales_by_product", vbInformation, "Sample Store"
