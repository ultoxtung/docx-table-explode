# docx-table-explode

TL;DR: This "script" iterates through tables in a docx file and removes them, while extracting and writing their data in the place of the old tables.

The day is 2069/04/20 (nice). After several attempts to exchange data between the US army and the aliens on the planet [REDACTED], the aliens send a large army to invade the Earth. With their superior technology, they effectively disarm all of the most large armies in the world. The hope for humanity now lies in the Neo Armstrong Cyclone Jet Armstrong Cannon, which was hidden under the Osaka Castle in Japan since the Edo period.  After the Sengoku family build the golden castle underneath the Osaka castle, the cannon was moved to under the golden castle.

Unfortunately, after the great downfall of the Sengoku family in 2012 and the disband of Omi Alliance in 2019, now no one alive knows exactly how to summon the golden castle up. After putting a lot of hard and honest work, the greatest minds alive on Earth finally successfully solved the riddle and bring the golden castle to the surface. But now there is another problem: How to bring the cannon out. The riddle is already there, but Japaneses being Japaneses, it is a docx file stored in a floppy disk. The rumour says if you write all the data in the table, cell by cell, in a single column, you will know the ritual to summon the Winged Dragon of Ra, which will be sacrified to special summon the cannon.

As the enemy is already at the verge of erase Osaka from the Earth, the humans only have barely enough time to do a quick search and find a FOSS software that can solve the problem. However, no one even bother doing this shit, since it is no practical reason doing this in the first place. In the mere last seconds before the last hope of humanity is destroyed, they can only send one last D-Mail (DeLorean Mail) back to a random point in history and hope the history will change.

That said D-Mail had stuck in my spam folder for multiple months now (nice try Google), but now the code is here. Maybe due to butterfly effect, doing this will prevent the war from breaking out in the first place, but whatever, I'm not sure if I will live to that point anyway.

# Usage

```
usage: main.py [-h] [--with-column-title] [--with-row-title] [--column-first] [--ignore-column IGNORE_COLUMN] [--ignore-row IGNORE_ROW] [--limit LIMIT] path

positional arguments:
  path                  Path to the file that need to be processed

options:
  -h, --help            show this help message and exit
  --with-column-title   Use this flag if table has first row used as column title
  --with-row-title      Use this flag if table has first column used as row title
  --column-first        Use this flag for exploding by column first then row, default is row before column
  --ignore-column IGNORE_COLUMN
                        Ignore these columns, separated by comma (count from 1)
  --ignore-row IGNORE_ROW
                        Ignore these rows, separated by comma (count from 1)
  --limit LIMIT         Only process the first n table(s)
```

# License

I would rather posting my driver's license online than writing a license for this thing.