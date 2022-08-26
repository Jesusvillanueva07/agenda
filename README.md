# agenda
registro para actividades
jesus david villanueva ficha: 2558427
---
## (hoja de vida)[https://github.com/Jesusvillanueva07/proyecto-hoda-de-vida.git]
## dia 12/08/22
### programacion visual basic
sub sena ()

 nom = "luis"

 msgbox nom

 num = 10

 msgbox num

 nom = "maria"

 msgbox "el nombre es:" & nom 
 
end sub

### taller visual basic 1
Sub programa()
  Z = InputBox("pago anual:")
  
  If Z < 1000 Then
    MsgBox (" no se realiza pago")
    
    
  Else
  
   If Z >= 1001 And Z < 10000 Then
    impuesto = 0.05
    MsgBox ("el pago anueal es:") & Z * impuesto
    
   Else
   
    
    If Z >= 10001 And Z < 100000 Then
     impuestos = 0.1
     MsgBox (" el pago anual es:") & Z * impuestos
     

     Else
     
     If Z >= 100001 And Z < 1000000 Then
       impuestos = 0.15
       MsgBox ("el pago anual es: ") & Z * impuestos
       
       
       Else
       
       If Z >= 1000001 And Z < 10000000 Then
       impuestos = 0.2
       MsgBox ("el pago anual es:") & Z * impuestos
       
       
       Else
       
       If Z >= 10000001 Then
       impuestos = 0.25
       MsgBox ("el pago anual es:") & Z * impuestos
     
     
     End If
     
           End If
     End If
     
           End If
     End If
     
           End If
     
     