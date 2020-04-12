# slidder

slid[i]der [the i is silent] aka slide image updater

*Does this ever happen to you?*

You are working on a presentation with 50 slides and 75 plots that you've generated,
but now the boss wants to change all the colors to purple.

With slidder you can easily update all the figures in your presentation with the push of a button!

slidder takes a Slides presentation and a local directory, and updates plots in the presentation
from files within that directory.

Matching is done based on writing the filename in the "alt text" description part
of the image object in the Slides presentation.

*NOTE:* for technical reasons, filenames (including path) CANNOT have spaces in them.

# TODOs

- [ ] proper installation
- [ ] example / readme on usage
- [ ] `add_image` -- insert an image to a given slide
- [ ] `id_images` -- find images in the presentation and figure out which files they are,
      annotating them in the Description field
- [ ] error handling of all GAPI requests
- [ ] LICENSE file


# Privacy policy

slidder needs read-write access to your Google Slides in order to do its job.
It also needs read-write access your Drive to find Slides presentations by name and to upload
your plots temporarily.

These data are only used during the duration of the command execution and are not abused.
