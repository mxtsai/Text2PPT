# Text2PPT Program
Generates a PPT file for worship use in less than 5 seconds.

Watch this short clip for how it works:

https://user-images.githubusercontent.com/31272872/174447465-18bfabce-86af-4bdd-8223-297249268893.mov


## Computer Setup
1. Download and install Python from here: https://www.python.org/downloads/
2. If on Windows Device, open up Command Prompt (type Windows Key + X), and pase the following code: `py -m pip install python-pptx`
3. If on Mac Device, open up Terminal and type: `pip3 install python-pptx`
4. Download this project as zip (click the green 'Code' button, and press 'Download ZIP'). And unzip the zip to a desktop folder. 
5. Double click `generate_ppt.py` to run the program.
6. If any questions, bring your PC to Maxwell and have him fix it.

## How to Use
Let's use an example! 
1. Prepare a worship song lyric: `永恆唯一的盼望.txt`

```
永恆唯一的盼望

有一位真神 祂名字叫耶穌
祂來到這世界上
為要救贖你
生命的意義
盡在這福音裡
只要你開口呼喊 耶穌

耶穌是生命
一切問題的解答
耶穌是生命
一切黑暗的亮光
將憂傷變為喜樂
將懼怕變為力量
耶穌是永恆唯一的盼望
```

When formatting the text file, keep song sections (e.g. verse, chorus, bridge) seperate by empty lines. 
Lyrics from different song sections will not be placed in the same slide.

2. Run `generate_ppt.py`, and drag the text file (i.e. `永恆唯一的盼望.txt`) into the program
3. Specify the number of lines each slide allows (try to use either 1 or 2)
4. Press Enter to run the program, and a PPT with the same filename will be saved beside the original text file (i.e. `永恆唯一的盼望.pptx`)

## Tips
1. For lyrics that require English and Chinese translation, place the English and Chinese translation in 2 consecutive lines.
When running the program, set the number of lines to print as '2' so that English and Chinese lyrics show up together.

For example: 
```
English worship text here
中文歌詞在這裡
```

As an example, we use the song `榮美的救主 Beautiful Saviour`

```
榮美的救主 Beautiful Saviour

耶穌
Jesus
榮美的救主 
Beautiful Saviour
復活全能真神
God of all majesty
榮耀君王
Risen King
神羔羊
Lamb of God
聖潔和公義
Holy and righteous
尊榮的拯救者
Blessed Redeemer
明亮晨星
Bright morning star


天使高聲齊頌揚
All the heavens shout Your praise
萬物都屈膝來敬拜你
All creation bows to worship You


何等奇妙
How wonderful
何等榮耀
How beautiful
超乎萬民之上
Name above every name
配得尊崇
Exalted high

何等奇妙
How wonderful
何等榮耀
How beautiful
耶穌我主
Jesus Your name
超乎萬名之上 耶穌
Name above every name, Jesus

我要永遠敬拜
I will sing forever
耶穌我愛你
Jesus I love You
耶穌我愛你
Jesus I love You
```
