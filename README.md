# power-slide

Power-Slide is a presentation tool written in PowerShell.

I wrote it for a talk I gave about PowerShell, so that I can demo what I am talking about and don't have to leva the shell but still have some minimal slides.

# Usage
Where you see ':somestring' it means typing ':somestring', where you see 'some_key' it means pressing 'some' key on the keyboard

:edit -> edit presentation

    pgdwn_key -> next slide
    
    pgup_key -> previous slide
    
    esc_key -> back to main menu
    
    insert_key -> insert new slide after the current one
    
    delete_key -> delete the current slide
    
    enter_key -> select this slide
    
      :title -> change title
    
      :subtitle -> change title
    
      :bullet -> write bullet points
    
      :done -> back to edit menu
    
:theme -> load theme file

:present -> start presentation

:save -> save the current slideshow to a csv file

:load -> load a slideshow from csv file

:q -> quit

After starting the presentation, you have some global functions and variables you can operate with

$_slides -> your slides as ps objects

$_selectedItem -> the index of the current slide


next -> next slide

prev -> prev slide

print -> print the current slide

done -> clean up all the exported functions and variables

# Themes
You can write your own theme for Power-Slide, you can find an example theme in this repository for that. Your theme should contain a single scriptblock that gets a slide object as a parameter (slide objects have Title, Subtitle and Points properties).

## Code injection
Note that your theme will be invoked when printing slides which makes it voulnerable for code injection.
