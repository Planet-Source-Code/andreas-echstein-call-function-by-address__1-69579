<div align="center">

## Call Function by Address


</div>

### Description

Call any Function by its address, API functions inside Dlls as well as your own Functions. Number of parameters is adjustable.

How to use: Just call CallByAddress with a pointer to the function and a parameter-array.

Example included.

How it works:

The windows api-function CallWindowProc will be used to call a small assembler-function which takes care of everithing for you.
 
### More Info
 
All inputs are of 4 bytes, representating a long or dword.

This comes in handy when you want to write Plugins for programs that are written in C/C++ and give you some kind of Callback-Address. With this code you can easily add the ability to call fuction pointers to VB.

The return value of the Function must be long, of course.


<span>             |<span>
---                |---
**Submitted On**   |2007-11-02 23:11:00
**By**             |[Andreas Echstein](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andreas-echstein.md)
**Level**          |Advanced
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Call\_Funct2089511142007\.zip](https://github.com/Planet-Source-Code/andreas-echstein-call-function-by-address__1-69579/archive/master.zip)








