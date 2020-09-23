Hi!!!
This is a Virtual Drive creator, I use a DOS command to do this, i'm sure there are others way to do this with out DOS (programs like VirtualCD do the samething without DOS i think), i should take a look at PSC i'm sure the answer is there but i have a bad memory and always forget, well if you know other way i'll thank you if mail me.

The DOS command i used is subst, the syntax is very simple:
SUBST [Drive1: [Drive2:]Path]
Example: SUBST X: C:\Temp
This will create a virtual drive (X:) of the folder C:\Temp

To remove a virtual drive:
SUBST /D Drive1:
Example: SUBST /D X:
This will delete the virtual drive.

The problem of using this method is that when you turn the computer off the virtual drive is gone. So you have to create it again, however if you make a BAT file and put it at windows start up, the drive is created automatically, i'm sure there are other ways to create virtual drives in windows, if you find one mail me!