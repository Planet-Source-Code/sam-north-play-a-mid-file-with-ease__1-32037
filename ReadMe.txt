MIDI

S North, Feb 2002

Any requests?

Yeah, vote for me! I can't afford this expensive software (leggally anyway ;)) because I am a very poor student - and in unison, say Ahh, didoms. tee hee hee. Thanks!

Thanks:

www.vbapi.com - for the function definitions on the MCI API and information on how to use it. Very useful site!

Why?

1. I built this app to demonstrate how to use the MCI sequencer to play a .mid file in Visual Basic. I was asked to make this by a mate, and so that's it evolved.

Purpose?

Plays a .mid file, like the music.mid file included.

How to use

use OpenMID to open a .mid file ready to be played, the file name must be an absolute filename, not just "music.mid" like the example, where the music.mid file needs to be move to C:\ The alias is the "handle" you will use to reference the file later, so make it meaningful, it's a srting by the way...

PlayMID plays the .mid file using the alias as its parameter as explained above and demonstrated in the example.

StopMID stops the .mid file using the alias, this does not reset the file to the begining, so bit like pausing it really!

CloseMID releases the file and should be used when you have finished with it or when the application closes, like in the example.

Thats it! Have fun and enjoy my code!