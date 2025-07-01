# Proyecto: Calculadora Financiera con Macros en Excel

Este proyecto fue desarrollado en VBA para automatizar cálculos financieros básicos como:

* Interés simple
* Valor futuro
* Valor presente
* Tasa de interés
* Número de periodos

Se utiliza un formulario con controles (TextBox y ComboBox) que cambian de acuerdo a la operación seleccionada. Cada operación está ligada a un botón que ejecuta código VBA para hacer los cálculos y mostrar resultados tanto en el formulario como en una hoja de Excel correspondiente.

---

## Interés Simple

**Lógica:**

* El usuario ingresa el Valor Presente y la Tasa de Interés.
* Se convierte la tasa a mensual según el periodo elegido.
* Se calcula el interés simple: `interesSimple = p * i`

**Código relevante:**

```vba
If cbxOperacion.Value = "Interés Simple" Then
  ConvertirTasaMensual
  p = txtVP.Value
  i = tasaMensual
  interesSimple = p * i

  Sheets("Interés").Select
  Range("A3").EntireRow.Insert
  Range("A3").Value = p
  Range("B3").Value = tasa
  Range("C3").Value = cbxPeriodoInteres.Value
  Range("D3").Value = interesSimple

  txtInteresSimple.Text = interesSimple
  txtInteresSimple.Value = Format(txtInteresSimple.Value, "Currency")
  txtTasaMensual.Text = tasaMensual * 100
  txtTasaMensual.Value = Format(txtTasaMensual.Value / 100, "0%")
End If
```

**Función ConvertirTasaMensual:**
Convierte tasas de interés desde su periodo original a una tasa mensual.

---

## Valor Futuro

**Lógica:**

* El usuario proporciona el Valor Presente, la Tasa de Interés y el número de periodos.
* Se convierte la tasa y el periodo a formato mensual.
* Se calcula el Valor Futuro: `VF = VP + intereses`, donde `intereses = VP * tasa * t`

**Código relevante:**

```vba
ElseIf cbxOperacion.Value = "Valor Futuro" Then
  ConvertirTasaMensual
  ConvertirPeriodoMensual
  ConvertirPeriodoTasaMes

  p = txtVP.Value
  i = txtInteres.Value / 100
  t = periodoMensual / periodoTasaMes
  intereses = p * i * t
  valorFuturo = p + intereses

  Sheets("Valor Futuro").Select
  Range("A3").EntireRow.Insert
  Range("A3").Value = p
  Range("B3").Value = i
  Range("C3").Value = cbxPeriodoInteres.Value
  Range("D3").Value = periodo
  Range("E3").Value = cbxPeriodos.Value
  Range("F3").Value = valorFuturo
  Range("G3").Value = intereses

  txtVF.Value = Format(valorFuturo, "Currency")
  txtInteresTotal.Value = Format(intereses, "Currency")
End If
```

---

## Valor Presente

**Lógica:**

* El usuario proporciona el Valor Futuro, la Tasa de Interés y el número de periodos.
* Se convierten las tasas y periodos a formato mensual.
* Se calcula: `VP = VF / (1 + (i * t))`

**Código relevante:**

```vba
ElseIf cbxOperacion.Value = "Valor Presente" Then
  ConvertirTasaMensual
  ConvertirPeriodoMensual
  ConvertirPeriodoTasaMes

  p = txtVF.Value
  i = txtInteres.Value / 100
  t = periodoMensual / periodoTasaMes
  valorPresente = p / (1 + (i * t))
  interesGenerado = p - valorPresente

  Sheets("Valor Presente").Select
  Range("A3").EntireRow.Insert
  Range("A3").Value = p
  Range("B3").Value = i
  Range("C3").Value = cbxPeriodoInteres.Value
  Range("D3").Value = periodo
  Range("E3").Value = cbxPeriodos.Value
  Range("F3").Value = valorPresente

  txtVP.Value = Format(valorPresente, "Currency")
  txtInteresTotal.Value = Format(interesGenerado, "Currency")
End If
```

---

## Tasa de Interés

**Lógica:**

* El usuario proporciona el Valor Presente, Valor Futuro y número de periodos.
* Se calcula: `tasaInteres = (VF - VP) / (VP * t)`

**Código relevante:**

```vba
ElseIf cbxOperacion.Value = "Tasa de interés" Then
  ConvertirPeriodoMensual
  ConvertirPeriodoTasaMes

  p = txtVP.Value
  vf = txtVF.Value
  t = periodoMensual
  tasaInteres = ((vf - p) / (p * t)) * 100

  Sheets("Tasa de Interés").Select
  Range("A3").EntireRow.Insert
  Range("A3").Value = p
  Range("D3").Value = vf
  Range("E3").Value = tasaInteres / 100

  txtInteres.Value = Format(tasaInteres / 100, "0.00%")
End If
```

---

## Número de Periodos

**Lógica:**

* El usuario proporciona el Valor Presente, Valor Futuro y la tasa.
* Se calcula: `n = (VF - VP) / (VP * tasaMensual)`

**Código relevante:**

```vba
ElseIf cbxOperacion.Value = "Número de periodos" Then
  ConvertirTasaMensual
  p = txtVP.Value
  vf = txtVF.Value
  i = tasaMensual
  intereses = vf - p
  nPeriodos = intereses / (p * i)

  Sheets("Número de Periodos").Select
  Range("A3").EntireRow.Insert
  Range("A3").Value = p
  Range("B3").Value = vf
  Range("C3").Value = i
  Range("D3").Value = cbxPeriodoInteres.Value
  Range("E3").Value = nPeriodos

  txtPeriodos.Text = nPeriodos
End If
```

---

## Notas adicionales

* Todos los resultados se presentan en formularios con formato adecuado (moneda, porcentaje, etc.).
* Se utiliza la selección en `cbxOperacion` para activar/desactivar los campos necesarios y dar formato visual (gris para campos inactivos, verde para resultados).

---

## Calculadora en funcionamiento
* Caso práctico: Mario tiene una deuda por $25,000.00 para pagar dentro de 18 quincenas. Si la operación fue pactada a una tasa de interés simple igual al 31.56% anual ¿Cuánto tendrá que pagar? ¿Cuántos intereses generó?

<img width="978" alt="Captura de pantalla 2025-07-01 a la(s) 2 13 19 a m" src="https://github.com/user-attachments/assets/ee1bc00f-1eab-4045-a2a0-2a62c107f50c" />

<img width="699" alt="Captura de pantalla 2025-07-01 a la(s) 2 11 16 a m" src="https://github.com/user-attachments/assets/0b65db27-11f7-4e76-adeb-09a9c252524a" />

<img width="1178" alt="Captura de pantalla 2025-07-01 a la(s) 2 12 22 a m" src="https://github.com/user-attachments/assets/3205ea92-7f99-4949-a9f7-ca1ea929b104" />



**Autor:** Mario Pérez Landeros  
**Tecnología usada:** VBA para Excel  
**Objetivo:** Automatización de cálculos financieros

