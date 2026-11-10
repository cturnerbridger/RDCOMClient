.onLoad <-
function(lib, pkg) {
 .COMInit()
}

.onUnload <-
function(libpath) {
 tryCatch(.Call("R_initCOM", FALSE, PACKAGE = "RDCOMClient"), error = function(e) NULL)
 library.dynam.unload("RDCOMClient", libpath)
}
