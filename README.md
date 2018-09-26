# MusicMagpiePricer
Provide a list of ISBNs to check and the script checks 3 websites that buy 2nd hand books for the best price.

Using selenium to control a Google Chrome browser this script visits
www.momox.co.uk
www.musicmagpie.co.uk/sell-books
www.webuybooks.co.uk

It enters each ISBN (from ISBNToProcess.txt) into the search bar and records the site that offers the best price. 
It then outputs the results as an .xlsx file
