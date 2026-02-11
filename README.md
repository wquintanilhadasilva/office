# Instalação do MSOffice

## Ativação
Tem duas formas pra tentar resolver isso.
Uma é desinstalando a chave:
Se o seu for 64bits, cola esse comando seguido de enter:

```
cd C:\Program Files\Microsoft Office\Office16
```

Cola esse seguido de enter:

```
cscript ospp.vbs /dstatus
```

Se aparecer os finais das chaves, cola esse, no lugar do "x" coloca a chave que aparece em seguida aperta enter.
Se tiver mais de uma refaz o último comando pra todas. OBS: deve estar desconectado da conta.


2º passo é esse.
Meu computador> disco local C:> arquivos de programas>Microsoft Office> Office 2016> Ospprearm> botão esquedo do mouse> executa como administrador.
Esse comando rearma a chave.
Se aparecer os finais das chaves, cola esse, no lugar do "x" coloca a chave que aparece em seguida aperta enter.
Se tiver mais de uma refaz o último comando pra todas. OBS: deve estar desconectado da conta.


Após instalação, vá à pasta do office (C:\Program Files\Microsoft Office\Office16) e use os comandos:

```
cd C:\Program Files\Microsoft Office\Office16
```

```
cscript ospp.vbs /dstatus
```

```
cscript ospp.vbs /setprt:1688
```

```
cscript ospp.vbs /unpkey:6F7TH >nul
```

```
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
```

```
cscript ospp.vbs /sethst:107.175.77.7
```

```
cscript ospp.vbs /act
```


```
cd C:\Users\Wedson-Oobj\Downloads\MSOffice
```

# Referências

* [https://www.youtube.com/watch?v=KtB1vlIuP2A](https://www.youtube.com/watch?v=KtB1vlIuP2A)

* [https://gist.github.com/devomman/7b7bb1d156a9099a10a7be4ad55e3002](https://gist.github.com/devomman/7b7bb1d156a9099a10a7be4ad55e3002)

* [https://msguides.com/office-2021](https://msguides.com/office-2021)

#Nova forma:

```
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
```

```
cd /d %ProgramFiles%\Microsoft Office\Office16
```

```
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
```

```
cscript ospp.vbs /setprt:1688
```

```
cscript ospp.vbs /unpkey:6F7TH >nul
```

```
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
```

```
cscript ospp.vbs /sethst:23.226.136.46
```

```
cscript ospp.vbs /act
```
