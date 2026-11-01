============================================================================
WE Library 1.0  -  LÉAME / README
============================================================================

  Español: primera parte.   English: second part.

ESPAÑOL
-------

Qué es
------
  WE ("Word-Excel") es una biblioteca para programadores de C# que crea y
  modifica documentos de Word y hojas de Excel sin necesidad de tener
  Microsoft Office instalado. No es un programa que se abra con doble clic:
  se añade como referencia a un proyecto de Visual Studio.
  Es gratuito y de código abierto. Este fichero explica cómo obtenerlo y por
  qué Windows puede mostrar un aviso al instalarlo, para que usted decida
  con toda la información.

Dos formas de obtenerlo: usted elige
------------------------------------
  A) El instalador ya preparado (WE-Library-v1.0-es.msi, en español; hay
     también versión en inglés y multilingüe). Es lo más cómodo.
  B) El código fuente, para compilarlo usted mismo con Visual Studio
     Community, que es gratuito.
  Las dos formas dan exactamente el mismo resultado.

Sitio oficial: descárguela solo de aquí
---------------------------------------
  El único sitio oficial de WE Library es su página en GitHub, la del autor:
     https://github.com/HectorFortun2024/WE-Library
  Descargue el instalador y el código fuente solo desde esa página. La
  licencia MIT permite que otras personas copien WE Library, así que pueden
  aparecer copias en otras webs. El autor no las revisa ni las controla, y
  cualquiera podría haberlas modificado para meter algo dañino. Si encuentra
  WE Library en otro sitio, no la use: vaya al enlace oficial.

  El autor no se hace responsable de lo que pueda ocurrir con una copia
  obtenida fuera del sitio oficial, ni de no seguir las indicaciones de este
  fichero. La huella digital (explicada más abajo) le permite comprobar en
  un momento que su fichero es el original.

Por qué Windows puede mostrar un aviso
--------------------------------------
  El instalador no lleva "firma digital". Esa firma es un certificado que se
  compra cada año y que solo identifica a quien publica el programa: no
  analiza ni garantiza lo que el programa hace. Como WE Library es gratuito,
  el autor ha preferido no pagarla.

  Por eso, al descargar o abrir el instalador, el navegador o Windows
  (SmartScreen) pueden decir que el fichero "no se descarga con frecuencia"
  o que es una "aplicación no reconocida". Es el aviso habitual para
  cualquier programa sin firma, por bueno que sea: no quiere decir que se
  haya encontrado nada malo en él.

  Si quiere comprobar que el fichero es exactamente el que publicó el autor,
  sin ningún cambio, compare su huella digital (se explica más abajo). Y si
  prefiere no ejecutar el instalador, tiene la forma B: el código completo
  está a su disposición.

A) Instalar con el instalador (.msi)
------------------------------------
  Hay tres versiones del instalador, según el idioma de sus ventanas: en
  español (-es), en inglés (-en) y multilingüe (-all). Elija una; instalan
  lo mismo.
  1. Descárguelo del enlace oficial:
     https://github.com/HectorFortun2024/WE-Library/releases/latest
  2. Si el navegador avisa de que el fichero no se descarga con frecuencia,
     elija la opción de conservarlo.
  3. Haga doble clic en el fichero. Si aparece la ventana azul
     "Windows protegió su PC", pulse "Más información" y después
     "Ejecutar de todas formas".
  4. Si Windows pregunta "¿Quiere permitir que esta aplicación haga cambios
     en el dispositivo?" con "Editor: desconocido", pulse "Sí".
     "Desconocido" solo significa que no lleva firma.
  5. Siga el asistente. Dónde queda cada cosa y cómo desinstalarlo se
     explica en "WE Library - Guía de Instalación.pdf" (carpeta
     "WE (DOCUMENTOS OFICIALES)").

  Comprobar la huella digital (opcional):
     Abra PowerShell en la carpeta donde descargó el fichero y escriba (con
     el nombre del que haya descargado):
        Get-FileHash "WE-Library-v1.0-es.msi"
     El valor que aparece en Hash debe ser exactamente el que corresponda:
     WE-Library-v1.0-es.msi:
     564AD461B30D73E72128A05E6DE0CF3AB96F1B7F4EB4FA1E8C07FD76CC76D97B
     WE-Library-v1.0-en.msi:
     2B0CF088A62BE78ED04A7101EABB654790DB9010F7E536EC61F15D60AEEFB0D2
     WE-Library-v1.0-all.msi:
     B9C13924E70FD2E5B04375C721A200947D7821F124B0FCDF1493DC5088D8A681
     Si no coincide, no lo instale y descárguelo de nuevo del enlace
     oficial.

B) Compilarlo usted mismo con Visual Studio Community
-----------------------------------------------------
  Así no ejecuta nada preparado por otra persona: la biblioteca la construye
  su propio Visual Studio a partir del código, que puede leer línea a línea
  antes de compilarlo.
  1. Instale Visual Studio Community (gratuito para particulares y para
     proyectos de código abierto). Durante la instalación marque
     "Desarrollo de escritorio de .NET". Se descarga de:
     https://visualstudio.microsoft.com/vs/community/
  2. Descargue el código: en la página del repositorio, botón verde "Code" y
     luego "Download ZIP". El repositorio está en:
     https://github.com/HectorFortun2024/WE-Library
  3. Antes de descomprimirlo: clic derecho sobre el ZIP, "Propiedades",
     marque "Desbloquear" y pulse "Aceptar". Así Visual Studio no tratará
     cada fichero como venido de Internet ni pondrá pegas al compilar.
  4. Descomprímalo y abra WE\WE.sln con Visual Studio. Si pregunta si confía
     en el proyecto, es su comprobación normal para todo lo que viene de
     Internet: conteste que sí.
  5. En la barra de arriba elija "Release" y pulse "Compilar" y luego
     "Compilar solución".
  6. La biblioteca queda en WE\bin\Release\net8.0-windows8.0\WE.dll .
     Añádala como referencia a su proyecto.

  Requisito: el SDK de .NET 8 o posterior (Visual Studio 2022 versión 17.8 o
  posterior, o Visual Studio 2026).

  La primera vez que compile hace falta conexión a Internet: Visual Studio
  descarga automáticamente los dos paquetes que usa WE, ClosedXML 0.102.2 y
  DocumentFormat.OpenXml 2.16.0 (ambos de licencia MIT). Hay dos programas
  de ejemplo: PruebaWord y PruebaExcel (también en PruebaWord.rar y
  PruebaExcel.rar).

Licencia
--------
  Licencia MIT: puede usarla, copiarla, modificarla y distribuirla
  libremente, también en programas comerciales. Los paquetes ClosedXML y
  DocumentFormat.OpenXml también son MIT. Texto completo en "WE (DOCUMENTOS
  OFICIALES)\End User Agreement.rtf".


============================================================================

ENGLISH
-------

What it is
----------
  WE ("Word-Excel") is a library for C# programmers that creates and
  modifies Word documents and Excel sheets without needing Microsoft Office
  installed. It is not a program you open with a double click: you add it as
  a reference to a Visual Studio project.
  It is free and open source. This file explains how to get it and why
  Windows may show a warning when you install it, so you can decide with all
  the facts.

Two ways to get it: your choice
-------------------------------
  A) The ready-made installer (WE-Library-v1.0-en.msi, in English; there are
     also Spanish and multilingual versions). The easiest way.
  B) The source code, to build it yourself with Visual Studio Community,
     which is free.
  Both ways give exactly the same result.

Official site: download it only from here
-----------------------------------------
  The only official site for WE Library is its GitHub page, the author's
  own:
     https://github.com/HectorFortun2024/WE-Library
  Download the installer and the source code only from that page. The MIT
  license allows other people to copy WE Library, so copies may appear on
  other websites. The author does not review or control them, and anyone
  could have modified them to add something harmful. If you find WE Library
  anywhere else, do not use it: go to the official link.

  The author is not responsible for anything that may happen with a copy
  obtained outside the official site, or from not following the instructions
  in this file. The digital fingerprint (explained below) lets you check in
  a moment that your file is the original.

Why Windows may show a warning
------------------------------
  The installer has no "digital signature". That signature is a certificate
  bought every year that only identifies who publishes the program: it does
  not analyze or guarantee what the program does. Since WE Library is free,
  the author chose not to pay for it.

  That is why, when you download or open the installer, the browser or
  Windows (SmartScreen) may say the file "isn't commonly downloaded" or is
  an "unrecognized app". This is the usual warning for any unsigned program,
  however good it is: it does not mean anything bad was found in it.

  If you want to check that the file is exactly the one the author
  published, unchanged, compare its digital fingerprint (explained below).
  And if you prefer not to run the installer, you have way B: the complete
  source code is available to you.

A) Install with the installer (.msi)
------------------------------------
  There are three versions of the installer, depending on the language of
  its windows: Spanish (-es), English (-en) and multilingual (-all). Choose
  one; they install the same thing.
  1. Download it from the official link:
     https://github.com/HectorFortun2024/WE-Library/releases/latest
  2. If the browser warns that the file isn't commonly downloaded, choose
     the option to keep it.
  3. Double-click the file. If the blue window "Windows protected your PC"
     appears, click "More info" and then "Run anyway".
  4. If Windows asks "Do you want to allow this app to make changes to your
     device?" with "Publisher: Unknown", click "Yes". "Unknown" only means
     it is not signed.
  5. Follow the wizard. Where everything goes and how to uninstall it is
     explained in "WE Library - Installation Guide.pdf" (folder
     "WE (DOCUMENTOS OFICIALES)").

  Checking the digital fingerprint (optional):
     Open PowerShell in the folder where you downloaded the file and type
     (with the name of the one you downloaded):
        Get-FileHash "WE-Library-v1.0-en.msi"
     The value shown under Hash must be exactly the matching one:
     WE-Library-v1.0-es.msi:
     564AD461B30D73E72128A05E6DE0CF3AB96F1B7F4EB4FA1E8C07FD76CC76D97B
     WE-Library-v1.0-en.msi:
     2B0CF088A62BE78ED04A7101EABB654790DB9010F7E536EC61F15D60AEEFB0D2
     WE-Library-v1.0-all.msi:
     B9C13924E70FD2E5B04375C721A200947D7821F124B0FCDF1493DC5088D8A681
     If it does not match, do not install it; download it again from the
     official link.

B) Build it yourself with Visual Studio Community
-------------------------------------------------
  This way you run nothing prepared by someone else: your own Visual Studio
  builds the library from the source code, which you can read line by line
  before building it.
  1. Install Visual Studio Community (free for individuals and for open
     source projects). During setup, tick ".NET desktop development".
     Download it from:
     https://visualstudio.microsoft.com/vs/community/
  2. Download the code: on the repository page, green "Code" button, then
     "Download ZIP". The repository is at:
     https://github.com/HectorFortun2024/WE-Library
  3. Before unzipping it: right-click the ZIP, "Properties", tick "Unblock"
     and click "OK". This way Visual Studio will not treat every file as
     coming from the Internet or object when building.
  4. Unzip it and open WE\WE.sln with Visual Studio. If it asks whether you
     trust the project, that is its normal check for anything coming from
     the Internet: answer yes.
  5. In the top bar choose "Release" and click "Build", then
     "Build Solution".
  6. The library is in WE\bin\Release\net8.0-windows8.0\WE.dll . Add it as a
     reference to your project.

  Requirement: the .NET 8 SDK or later (Visual Studio 2022 version 17.8 or
  later, or Visual Studio 2026).

  The first build needs an Internet connection: Visual Studio automatically
  downloads the two packages WE uses, ClosedXML 0.102.2 and
  DocumentFormat.OpenXml 2.16.0 (both MIT licensed). There are two sample
  programs: PruebaWord and PruebaExcel (also in PruebaWord.rar and
  PruebaExcel.rar).

License
-------
  MIT license: you may use, copy, modify and distribute it freely, including
  in commercial programs. The ClosedXML and DocumentFormat.OpenXml packages
  are MIT too. Full text in "WE (DOCUMENTOS OFICIALES)\End User
  Agreement.rtf".


----------------------------------------------------------------------------
Autor / Author: Héctor Fortún Castañón
