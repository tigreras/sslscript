import subprocess as sp
import win32com.client as win32
import os
outlook = win32.Dispatch('outlook.application')


#path = 'C:/Users/U072749/Downloads/OpenSSL-Win64/OpenSSL-Win64/bin'
# if os.path.exists("ca-cert"):
#     sp.call(["rm", "-rf", "ca-cert"])
#     sp.call(["mkdir", "ca-cert"])

# 0. Edit File SAN.CNF
# 1. Sesuaikan DNS1 dan DNS2 Dengan Request

open ("san.cnf", "w").write(
"""
[ req ]
default_bits= 2048
distinguished_name= req_distinguished_name
req_extensions= req_ext
[ req_distinguished_name ]
countryName		= Country Name
stateOrProvinceName	= State
localityName		= Locality
organizationName	= Organization
organizationalUnitName	= Organizational Unit
commonName		= Common Name
[ req_ext ]
subjectAltName = @alt_names
[alt_names]

DNS.1 =  tes.intra.bca.co.id
DNS.2 =  tes.intra.bca.co.id

"""
)

# 2. Nama File CSR dan Key Dapat Diedit Disini
sp.call(['openssl', 'req', '-out', 'tes.intra.bca.co.id.csr','-newkey', 'rsa:2048', '-nodes', '-keyout', 'tes.intra.bca.co.id.key', '-config', 'san.cnf'])

# 3. Kirim Email



mail = outlook.CreateItem(0)
print('Masukkan Subject Email')
subject = input()
mail.Subject = subject
print('Masukkan Resipien Email') 
to = input()
mail.To = to
print('Masukkan CC Email') 
cc = input()
mail.CC = cc
mail.HTMLBody = r"""
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:#1F497D;">Dear Tim SSL Support,</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:#1F497D;">&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:#1F497D;">Mohon bantuannya utk sign CSR terlampir terkait project OCP dengan subjek diatas.</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:#1F497D;">Algoritma yg digunakan sha256 dgn CA Prod.</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:#1F497D;">&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:#2F5496;">Terima kasih atas bantuan &amp; perhatiannya.</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:#2F5496;">&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:#2F5496;">Best Regards,</span></p>
<p style='margin:0cm;margin-bottom:12.0pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:#2F5496;">Leonardo Wijaya.</span></p>
"""
mail.Attachments.Add(os.getcwd() + "\\tes.intra.bca.co.id.csr")
mail.Send()


