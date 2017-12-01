Computer Lab
============
I've managed a couple computer labs before. In one case I had Windows
Server and in another I didn't. In either case, you can batch install all
the software you want pretty easily. See *install_SSL.bat* and
*install_YapSDA.bat* for two different examples.

At one school, I also wanted to do a system restore automatically on
shutdown sometimes. In Windows Server I'd set the system restore create
script to run *create_system_restore.bat* (and then disable it so it
doesn't keep creating system restore points) and then enable
*system_restore.vbs* on shutdown.

See *notes.html* for some additional notes about configuring stuff.
