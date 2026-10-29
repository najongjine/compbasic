import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const dir = path.dirname(fileURLToPath(import.meta.url));
const songs = {};
// key | artist | title | verified YouTube video ID (otherwise an explicit search link)
const catalogue = `
trot|장윤정|어머나
balladkr|박효신|야생화
folkkr|김광석|서른 즈음에
rockkr|YB|나는 나비
hipkr|다이나믹 듀오|출첵
rbkr|DEAN|D (half moon)
indiekr|잔나비|주저하는 연인들을 위해
arirang|국립국악원|아리랑
pansori|안숙선|춘향가 사랑가
dynamite|BTS|Dynamite|gdZLi9oWNZg
adele|Adele|Someone Like You|hLQl3WQQoQ0
acdc|AC/DC|Back In Black|pAgnJDJN4VA
metal|Metallica|Master of Puppets
jazz|The Dave Brubeck Quartet|Take Five
blues|B.B. King|The Thrill Is Gone
alicia|Alicia Keys|If I Ain't Got You|Ju8Hr50Ckwk
aretha|Aretha Franklin|Respect
uptown|Mark Ronson feat. Bruno Mars|Uptown Funk|OPf0YbXqDm0
disco|Bee Gees|Stayin' Alive
hiphop|Dr. Dre feat. Snoop Dogg|Still D.R.E.
rap|Eminem|Lose Yourself
reggae|Bob Marley & The Wailers|Three Little Birds
country|John Denver|Take Me Home, Country Roads
folk|Bob Dylan|Blowin' in the Wind
acoustic|Jason Mraz|I'm Yours
carpenters|Carpenters|(They Long to Be) Close to You
eagles|Eagles|Hotel California
easy|Norah Jones|Don't Know Why
bossa|Stan Getz & João Gilberto feat. Astrud Gilberto|The Girl from Ipanema
lofi|Nujabes|Aruarian Dance
ambient|Brian Eno|An Ending (Ascent)
newage|Yiruma|River Flows in You
chill|Air|La femme d'argent
edm|Avicii|Wake Me Up
house|Daft Punk|Around the World
techno|Jeff Mills|The Bells
trance|Robert Miles|Children
euro|2 Unlimited|No Limit
takeonme|a-ha|Take On Me|djV11Xbc914
electro|Lady Gaga|Poker Face
dnb|Roni Size / Reprazent|Brown Paper Bag
beethoven|Ludwig van Beethoven|Symphony No. 5 · 1st movement
dvorak|hr-Sinfonieorchester · Andrés Orozco-Estrada|Dvořák Symphony No. 9|jOofzffyDSA
opera|Luciano Pavarotti|Nessun dorma
chamber|Franz Schubert|String Quartet No. 14 · Death and the Maiden
time|Hans Zimmer|Time · Live in Prague|va1oiojnGrA
epic|Two Steps From Hell|Heart of Courage
musical|Les Misérables Cast|One Day More
happy|Pharrell Williams|Happy|ZbZSe6N_BXs
queen|Queen|Don't Stop Me Now|HgzGwKwLmgM
playful|Henry Mancini|The Pink Panther Theme
celebrate|Kool & The Gang|Celebration
christmas|Mariah Carey|All I Want for Christmas Is You
fixyou|Coldplay|Fix You
melancholy|Radiohead|No Surprises
lonely|김광석|어느 60대 노부부 이야기
bittersweet|ABBA|The Winner Takes It All
mournful|Samuel Barber|Adagio for Strings
tragic|Pyotr Ilyich Tchaikovsky|Symphony No. 6 · 4th movement
gymnopedie|Erik Satie|Gymnopédie No. 1
imagine|John Lennon|Imagine|YkgkThdzX-8
lullaby|Johannes Brahms|Wiegenlied · Op. 49 No. 4
clair|Claude Debussy|Clair de lune
dreamy|Cocteau Twins|Heaven or Las Vegas
allofme|John Legend|All of Me|450p7goxZqg
thousand|Christina Perri|A Thousand Years|rtOvBOTyX00
lovely|Stevie Wonder|Isn't She Lovely
love|Elvis Presley|Can't Help Falling in Love
passion|Édith Piaf|La Vie en rose
hope|Louis Armstrong|What a Wonderful World
nostalgia|산울림|청춘
superman|John Williams|Superman March
starwars|John Williams|Star Wars Main Title
mars|Gustav Holst|The Planets · Mars
rocky|Bill Conti|Gonna Fly Now
wewill|Queen|We Will Rock You
numb|Linkin Park|Numb|kXYiU_JCYtU
badguy|Billie Eilish|bad guy|DyDfgMOUjCI
teardrop|Massive Attack|Teardrop
jaws|John Williams|Jaws Theme
halloween|John Carpenter|Halloween Theme
psycho|Bernard Herrmann|Psycho · The Murder
requiem|Wolfgang Amadeus Mozart|Requiem · Lacrimosa
creep|Radiohead|Climbing Up the Walls
retro|The Weeknd|Blinding Lights|4NRXx6U8ABQ
vintage|The Chordettes|Mr. Sandman
reflect|Simon & Garfunkel|The Sound of Silence
wish|이문세|옛사랑
timeless|The Beatles|Yesterday
whiskey|Chris Stapleton|Tennessee Whiskey|4zAThXFOy2c
billie|Michael Jackson|Billie Jean|Zi_XLOBDo_Y
shape|Ed Sheeran|Shape of You|JGwWNGJdvx8
getlucky|Daft Punk feat. Pharrell Williams & Nile Rodgers|Get Lucky|5NV6Rdv1a3I
fast|DragonForce|Through the Fire and Flames
cantstop|Red Hot Chili Peppers|Can't Stop
duet|이소라 · 김현철|그대 안의 블루
choir|The Tabernacle Choir|Hallelujah Chorus · Handel Messiah
backing|The Beatles|Because
cash|Johnny Cash|Hurt
bassvoice|Paul Robeson|Ol' Man River
baritone|Dietrich Fischer-Dieskau|Schubert · Erlkönig
alto|Kathleen Ferrier|Blow the Wind Southerly
mezzo|Elīna Garanča|Carmen · Habanera
soprano|Diana Damrau|The Magic Flute · Queen of the Night aria
falsetto|Bee Gees|How Deep Is Your Love
soft|아이유|밤편지
gentle|성시경|너의 모든 순간
sweet|AKMU|Give Love
smooth|Nat King Cole|L-O-V-E
breathy|Billie Eilish|when the party's over
airy|백예린|Square (2017)
whisper|Billie Eilish|everything i wanted
power|Whitney Houston|I Will Always Love You
husky|김건모|아름다운 이별
raspy|Rod Stewart|Maggie May
gritty|Janis Joplin|Piece of My Heart
raw|Nirvana|Where Did You Sleep Last Night · MTV Unplugged
shout|The Beatles|Twist and Shout
scream|Linkin Park|Given Up
vibrato|조수미|가곡 가고파
melisma|Mariah Carey|Vision of Love
legato|Franz Schubert|Ave Maria
staccato|Wolfgang Amadeus Mozart|The Magic Flute · Der Hölle Rache · Diana Damrau
spoken|Gil Scott-Heron|The Revolution Will Not Be Televised
chant|Gregorian chant|Dies irae
youth|NewJeans|Hype Boy
child|Disney · Katie Lopez / Agatha Lee Monn / Kristen Bell|Frozen · Do You Want to Build a Snowman · young Anna
retrovoice|박진영|When We Disco · duet with 선미
clear|아이유|마음을 드려요
bright|볼빨간사춘기|여행
darkvoice|Leonard Cohen|You Want It Darker
piano|Yiruma|Kiss the Rain
grand|Frédéric Chopin|Ballade No. 1 · Krystian Zimerman
epiano|Herbie Hancock|Chameleon
keyboard|Europe|The Final Countdown
organ|Johann Sebastian Bach|Toccata and Fugue in D minor BWV 565 · organ
hammond|Booker T. & the M.G.'s|Green Onions
accordion|Astor Piazzolla|Libertango · accordion Richard Galliano
harpsichord|Johann Sebastian Bach|Italian Concerto BWV 971 · harpsichord
guitar|Tommy Emmanuel|Classical Gas
electric|Guns N' Roses|Sweet Child O' Mine
classicalguitar|Francisco Tárrega|Recuerdos de la Alhambra · classical guitar
bassguitar|Jaco Pastorius|Portrait of Tracy
violin|Antonio Vivaldi|The Four Seasons · Summer · 3rd movement
viola|Béla Bartók|Viola Concerto · Tabea Zimmermann
cello|Johann Sebastian Bach|Cello Suite No. 1 · Prelude · Yo-Yo Ma
doublebass|Giovanni Bottesini|Elegy No. 1 · double bass
harp|Claude Debussy|Première Arabesque · harp · Lavinia Meijer
ukulele|Israel Kamakawiwoʻole|Somewhere Over the Rainbow
mandolin|Chris Thile|Bach Sonata No. 1 · mandolin
banjo|Earl Scruggs|Foggy Mountain Breakdown
drums|Dave Brubeck Quartet · Joe Morello|Take Five · drum solo
edrums|Phil Collins|In the Air Tonight
kick|Daft Punk|Revolution 909
snare|Maurice Ravel|Boléro · snare drum
hihat|Toto|Rosanna · Jeff Porcaro
cymbals|John Philip Sousa|The Stars and Stripes Forever
tambourine|The Byrds|Mr. Tambourine Man
bongo|Preston Epps|Bongo Rock
conga|Santana|Oye Como Va
timpani|Richard Strauss|Also sprach Zarathustra · opening
xylophone|George Hamilton Green|Ragtime Robin · xylophone
marimba|Keiko Abe|Dream of the Cherry Blossoms
flute|Claude Debussy|Syrinx · Emmanuel Pahud
piccolo|John Philip Sousa|The Stars and Stripes Forever · piccolo solo
clarinet|Wolfgang Amadeus Mozart|Clarinet Concerto K.622 · Adagio
oboe|Alessandro Marcello|Oboe Concerto in D minor · Adagio
bassoon|Wolfgang Amadeus Mozart|Bassoon Concerto K.191
sax|Gerry Rafferty|Baker Street
trumpet|Miles Davis|So What
trombone|Tommy Dorsey|I'm Getting Sentimental Over You
horn|Wolfgang Amadeus Mozart|Horn Concerto No. 4 K.495
tuba|Ralph Vaughan Williams|Tuba Concerto
harmonica|Stevie Wonder|Isn't She Lovely · harmonica solo
recorder|Antonio Vivaldi|Recorder Concerto RV 443 · Michala Petri
panflute|Gheorghe Zamfir|The Lonely Shepherd
gayageum|국립국악원|가야금산조
geomungo|국립국악원|거문고산조
haegeum|강은일|해금 · 비에 젖은 해금
daegeum|국립국악원|대금산조
piri|국립국악원|상령산 · 피리 독주
janggu|국립국악원|설장구
buk|국립국악원|판소리 고법 · 북 반주
kkwaenggwari|국립국악원|사물놀이 · 꽹과리
jing|국립국악원|사물놀이 · 징
taepyeongso|국립국악원|태평소 시나위
danso|국립국악원|청성곡 · 단소 독주
synthbass|New Order|Blue Monday
synthpad|Vangelis|Blade Runner Blues
machine|Kraftwerk|The Robots
bass808|Kanye West|Love Lockdown
arpeggio|Donna Summer|I Feel Love
stringmachine|Jean-Michel Jarre|Oxygène Part IV
strings|Pyotr Ilyich Tchaikovsky|Serenade for Strings · 1st movement
brass|Aaron Copland|Fanfare for the Common Man
woodwind|Claude Debussy|Petite Suite · woodwind quintet
orchestralperc|John Williams|Star Wars · The Imperial March
bohemian|Queen|Bohemian Rhapsody
november|Guns N' Roses|November Rain|8SbUC-UaAxE
build|Swedish House Mafia|Don't You Worry Child
drop|Martin Garrix|Animals
fade|The Beatles|Hey Jude
postchorus|Taylor Swift|Shake It Off
interlude|Michael Jackson|Thriller
seeyou|Wiz Khalifa feat. Charlie Puth|See You Again|RgKAFK5djSk
`;
for (const line of catalogue.trim().split('\n')) {
  const [key, artist, title, video] = line.split('|');
  songs[key] = { artist, title, ...(video ? {video} : {}) };
}
const mappings = new Map();
function group(category, subgroup, text) {
  mappings.set(category + '/' + subgroup, text.trim().split('\n').map(l=>{
    const [song, note] = l.split('|');
    if (!songs[song] || !note) throw new Error('Invalid example: '+l);
    return {song,note};
  }));
}
group('장르','한국 음악 장르',`
trot|꺾는 창법과 규칙적으로 출렁이는 반주를 들어보세요.
balladkr|조용한 시작에서 큰 후렴으로 넓어지는 한국 발라드의 흐름입니다.
dynamite|그룹 보컬과 정교한 팝 편곡을 함께 들어보세요.
folkkr|통기타와 이야기를 전하는 보컬이 중심입니다.
rockkr|밴드의 기타·드럼과 힘 있는 후렴에 집중하세요.
hipkr|한국어 랩의 리듬과 비트가 맞물리는 방식입니다.
rbkr|느긋한 그루브와 섬세한 보컬의 리듬을 들어보세요.
indiekr|밴드 사운드와 개성 있는 보컬의 질감을 비교하세요.
arirang|한국 전통 선율과 장단의 흐름을 들어보세요.
arirang|국악 연주의 음색·장단을 살펴보는 입문 예시입니다.
pansori|소리꾼의 표현과 북 반주가 주고받는 방식입니다.
dynamite|경쾌한 댄스 비트와 반복되는 후렴을 들어보세요.
`);
group('장르','대중적인 음악 장르',`
shape|쉽게 기억되는 선율과 반복되는 팝 곡의 구성입니다.
adele|피아노와 보컬을 중심으로 감정을 쌓는 발라드입니다.
eagles|일렉트릭 기타와 드럼·베이스의 밴드 소리를 들어보세요.
acdc|굵직한 기타 리프와 강한 드럼이 중심입니다.
metal|빠르고 촘촘한 기타 리프와 무거운 사운드입니다.
jazz|악기들의 리듬과 즉흥적인 솔로를 들어보세요.
blues|기타의 꺾는 음과 보컬의 응답을 들어보세요.
alicia|보컬의 리듬과 소울을 담은 피아노 반주입니다.
aretha|힘 있는 감정 표현과 풍부한 보컬의 울림입니다.
uptown|잘게 끊는 기타·베이스와 리듬의 탄력을 느껴보세요.
disco|일정한 댄스 비트와 반짝이는 반주가 특징입니다.
hiphop|반복되는 비트와 랩의 그루브가 중심입니다.
rap|말의 강세가 박자에 실리는 방식을 들어보세요.
reggae|엇박에 놓이는 기타와 느긋한 리듬을 들어보세요.
country|통기타 반주와 이야기를 들려주는 보컬입니다.
folk|간결한 반주 위에 노랫말과 선율이 앞에 나옵니다.
dynamite|춤추기 좋은 박자에 팝 후렴을 결합한 예시입니다.
`);
group('장르','조용하고 감성적인 음악',`
acoustic|통기타의 줄 소리와 가벼운 반주를 들어보세요.
adele|피아노가 노래의 화음과 흐름을 이끌어갑니다.
carpenters|강한 자극보다 부드러운 선율과 보컬이 중심입니다.
eagles|록 밴드의 소리를 비교적 여유롭게 풀어낸 예시입니다.
easy|편안한 리듬과 부담 없이 흐르는 보컬입니다.
bossa|부드러운 노래와 살짝 엇갈리는 기타 리듬입니다.
lofi|차분한 반복 비트와 부드럽게 번지는 질감입니다.
ambient|뚜렷한 드럼보다 길게 머무는 배경 음색을 들어보세요.
newage|피아노의 반복 패턴과 감성적인 선율입니다.
chill|편안한 베이스와 전자음이 만드는 휴식의 느낌입니다.
`);
group('장르','신나는 음악',`
edm|고조되는 흐름 뒤에 터지는 댄스 비트를 들어보세요.
house|네 박을 꾸준히 짚는 킥과 반복 그루브입니다.
techno|반복적인 전자 타악기와 점진적인 변화를 들어보세요.
trance|반복되는 신스 선율과 길게 고조되는 흐름입니다.
euro|빠른 전자 비트와 직관적인 후렴이 특징입니다.
dynamite|팝 선율과 경쾌한 댄스 리듬의 결합입니다.
takeonme|신시사이저가 이끄는 밝은 팝 선율입니다.
disco|일정한 킥과 춤추기 좋은 리듬을 들어보세요.
electro|전자음과 리듬이 팝 곡의 중심을 이룹니다.
dnb|빠르고 쪼개지는 드럼과 낮은 베이스를 들어보세요.
`);
group('장르','클래식 · 영화 · 무대',`
beethoven|짧은 동기가 발전하는 클래식 작품의 예시입니다.
dvorak|현악·관악·타악이 함께 만드는 합주를 들어보세요.
dvorak|교향곡의 큰 구성과 악기군의 대비를 들어보세요.
opera|극 속 인물을 노래하는 성악과 관현악 반주입니다.
chamber|소수의 현악기가 서로 주고받는 대화를 들어보세요.
time|영화 장면의 감정을 이끄는 음악의 예시입니다.
time|층층이 커지는 사운드로 영화 같은 공간감을 만듭니다.
epic|현악·금관·타악이 함께 만드는 큰 규모의 느낌입니다.
musical|인물들의 이야기가 노래와 합창으로 전개됩니다.
`);
group('분위기','밝고 신나는 분위기',`
happy|가벼운 리듬과 밝은 후렴이 주는 행복감입니다.
lovely|들뜬 멜로디와 보컬의 기쁨을 느껴보세요.
sweet|톡톡 튀는 리듬과 명랑한 보컬입니다.
dynamite|박자에 맞춰 몸이 움직이는 경쾌함입니다.
queen|쉬지 않고 앞으로 나아가는 에너지를 들어보세요.
edm|고조된 뒤 펼쳐지는 비트가 기대감을 만듭니다.
playful|장난스럽게 움직이는 선율과 리듬입니다.
uptown|리듬을 주고받는 재미와 흥겨움을 느껴보세요.
bright|밝은 보컬과 활기찬 반주의 조합입니다.
rockkr|후렴이 펼쳐지며 기분을 끌어올리는 느낌입니다.
christmas|종소리와 큰 편곡이 만드는 축제 분위기입니다.
celebrate|함께 축하하는 듯한 보컬과 그루브입니다.
`);
group('분위기','슬프고 감성적인 분위기',`
adele|이별의 감정이 드러나는 보컬을 들어보세요.
balladkr|절제된 시작에서 감정이 크게 고조됩니다.
melancholy|담담한 목소리와 쓸쓸한 반주의 대비입니다.
adele|후렴에서 드러나는 가슴 아픈 감정 표현입니다.
mournful|길게 이어지는 현악이 깊은 슬픔을 만듭니다.
lonely|소박한 반주와 회상하는 목소리의 외로움입니다.
wish|추억을 되짚는 선율과 감상적인 보컬입니다.
bittersweet|밝게 들리는 흐름 속에 씁쓸한 감정이 겹칩니다.
mournful|현악기의 긴 호흡이 애도의 정서를 만듭니다.
tragic|마지막 악장의 무겁게 가라앉는 흐름입니다.
seeyou|상실과 그리움을 담은 후렴의 정서입니다.
requiem|무거운 합창과 반주가 만드는 침울한 분위기입니다.
`);
group('분위기','편안하고 차분한 분위기',`
gymnopedie|성긴 피아노 음과 긴 여백을 들어보세요.
imagine|담담한 피아노와 목소리가 만드는 평온함입니다.
bossa|느긋하게 흐르는 리듬에 집중해보세요.
gentle|큰 자극 없이 이어지는 온화한 목소리입니다.
carpenters|부드러운 보컬과 둥근 반주의 조합입니다.
lullaby|반복되는 선율과 자장가의 안정감입니다.
clair|잔잔하게 번지는 피아노 음색입니다.
ambient|서두르지 않고 오래 머무는 소리입니다.
hope|목소리의 따뜻한 질감과 낙관적인 선율입니다.
easy|가까이서 노래하는 듯한 포근한 소리입니다.
dreamy|번지는 기타와 보컬이 만드는 몽환적인 질감입니다.
chill|여유로운 베이스와 느긋한 반복을 들어보세요.
`);
group('분위기','사랑스럽고 낭만적인 분위기',`
allofme|피아노와 보컬이 만드는 낭만적인 분위기입니다.
sweet|가볍고 달콤하게 표현하는 보컬입니다.
lovely|애정을 담은 목소리와 들뜬 선율입니다.
love|작은 소리로 다정하게 건네는 듯한 노래입니다.
passion|보컬의 큰 감정과 선율의 상승을 들어보세요.
hope|소박한 반주와 따뜻한 시선을 담은 노래입니다.
easy|목소리가 가까이 들리는 작은 공간의 느낌입니다.
thousand|부드러운 목소리와 이어지는 선율의 애정입니다.
fixyou|작게 시작해 함께 노래하는 듯 커지는 희망입니다.
nostalgia|지난 시절을 떠올리게 하는 멜로디입니다.
`);
group('분위기','웅장하고 강렬한 분위기',`
epic|크게 펼쳐지는 오케스트라와 타악의 규모감입니다.
power|후렴에서 풍부하게 터지는 목소리의 힘입니다.
november|조용한 구간과 큰 클라이맥스의 극적인 대비입니다.
starwars|금관의 넓은 선율이 만드는 장엄함입니다.
dvorak|여러 악기군이 함께 쌓아 올리는 큰 울림입니다.
metal|촘촘한 기타 리프와 강한 타격감을 들어보세요.
rocky|앞으로 나아가는 리듬과 승리의 분위기입니다.
superman|상승하는 금관 선율이 영웅적인 인상을 만듭니다.
wewill|간결하지만 자신감 있게 박자를 짚는 표현입니다.
numb|거친 기타와 강하게 밀어붙이는 보컬입니다.
scream|강한 드럼과 폭발하는 보컬의 에너지입니다.
time|소리가 점차 커지며 한 장면처럼 펼쳐집니다.
`);
group('분위기','어둡고 긴장감 있는 분위기',`
badguy|낮고 간결한 베이스가 만드는 어두운 질감입니다.
teardrop|반복되는 소리와 공기감이 신비롭게 느껴집니다.
jaws|짧은 동기가 반복되며 다가오는 긴장감입니다.
psycho|날카로운 현악의 반복이 만드는 팽팽한 느낌입니다.
halloween|반복되는 건반 패턴이 만드는 으스스함입니다.
teardrop|기억에 남는 음색과 불안한 듯한 여운입니다.
mars|낮고 거친 리듬과 다가오는 듯한 위협감입니다.
jaws|단순한 저음 동기가 불길한 예감을 만듭니다.
requiem|가라앉은 합창과 화음이 만드는 음울함입니다.
halloween|작은 패턴의 집요한 반복이 소름을 돋웁니다.
creep|불안정하게 겹치는 소리와 압박감입니다.
psycho|빠른 반복과 날카로운 소리의 불안함입니다.
`);
group('분위기','옛날 감성과 추억의 분위기',`
wish|옛 기억을 떠올리는 선율과 목소리입니다.
retro|1980년대풍 신스와 드럼의 질감을 들어보세요.
vintage|옛 녹음의 분위기와 보컬 하모니입니다.
smooth|고전적인 반주와 우아하게 부르는 보컬입니다.
nostalgia|지나간 시간을 돌아보는 노래의 정서입니다.
hope|포근한 목소리와 친근한 반주의 조합입니다.
bittersweet|아름다운 멜로디에 씁쓸한 감정이 겹칩니다.
reflect|담담하게 돌아보는 듯한 보컬과 여백입니다.
wish|멀어진 기억을 그리는 듯한 아련함입니다.
timeless|간결한 편곡과 오래 남는 선율의 예시입니다.
`);
group('속도','느린 속도',`
gymnopedie|음 사이의 여백과 느린 호흡을 들어보세요.
whiskey|서두르지 않는 반주와 길게 늘이는 보컬입니다.
adele|느린 흐름 위에 놓이는 피아노와 노래입니다.
easy|여유로운 반주와 가볍게 밀고 당기는 보컬입니다.
bossa|빠르게 몰아가지 않는 느긋한 리듬입니다.
ambient|뚜렷한 박자 없이 천천히 머무는 예시입니다.
`);
group('속도','보통 속도',`
shape|걷는 듯 편안하게 이어지는 중간 속도의 느낌입니다.
billie|지나치게 느리거나 빠르지 않은 리듬입니다.
rbkr|중간 템포 안에서 보컬이 여유 있게 움직입니다.
house|처음부터 끝까지 일정하게 반복되는 비트입니다.
getlucky|가볍게 몸을 움직이기 좋은 경쾌한 흐름입니다.
`);
group('속도','빠른 속도',`
queen|앞으로 달리는 듯한 빠른 피아노와 리듬입니다.
takeonme|빠른 신스 패턴과 활기찬 드럼입니다.
fast|촘촘하고 매우 빠르게 진행되는 기타·드럼입니다.
dynamite|경쾌한 댄스 팝의 빠른 체감을 비교하세요.
bright|산뜻하게 앞으로 나아가는 반주입니다.
dnb|잘게 나뉜 드럼이 매우 빠르게 움직입니다.
fast|고속으로 질주하는 리프와 타격감입니다.
`);
// Exact BPM is supplied by the metronome; song references illustrate pacing, not measured BPM.
group('속도','BPM으로 속도 지정하기',`
gymnopedie|느린 호흡의 참고곡입니다. 정확한 60 BPM은 아래 박자 듣기로 비교하세요.
adele|느린 발라드의 참고곡입니다. 정확한 70 BPM은 아래 박자 듣기로 비교하세요.
easy|느긋한 노래의 참고곡입니다. 정확한 80 BPM은 아래 박자 듣기로 비교하세요.
rbkr|여유로운 그루브의 참고곡입니다. 정확한 90 BPM은 아래 박자 듣기로 비교하세요.
shape|중간 템포 느낌의 참고곡입니다. 정확한 100 BPM은 아래 박자 듣기로 비교하세요.
billie|경쾌한 그루브의 참고곡입니다. 정확한 110 BPM은 아래 박자 듣기로 비교하세요.
getlucky|댄스 리듬의 참고곡입니다. 정확한 120 BPM은 아래 박자 듣기로 비교하세요.
house|반복되는 댄스 비트의 참고곡입니다. 정확한 130 BPM은 아래 박자 듣기로 비교하세요.
trance|빠른 전자 음악의 참고곡입니다. 정확한 140 BPM은 아래 박자 듣기로 비교하세요.
queen|활기찬 빠르기의 참고곡입니다. 정확한 150 BPM은 아래 박자 듣기로 비교하세요.
takeonme|빠르게 진행되는 팝의 참고곡입니다. 정확한 160 BPM은 아래 박자 듣기로 비교하세요.
dnb|고속 드럼의 참고곡입니다. 정확한 180 BPM은 아래 박자 듣기로 비교하세요.
`);
group('목소리','남성·여성 목소리',`
gentle|남성 보컬의 음색과 프레이징을 들어보세요.
soft|여성 보컬의 음색과 프레이징을 들어보세요.
allofme|남성 가수의 솔로 보컬 예시입니다.
adele|여성 가수의 솔로 보컬 예시입니다.
duet|두 목소리가 번갈아 부르고 만나는 부분입니다.
duet|남녀 보컬의 음색과 화음의 대비입니다.
dynamite|여러 멤버가 주고받는 보컬을 들어보세요.
choir|여러 성부가 함께 만드는 큰 목소리의 울림입니다.
backing|주선율을 감싸는 겹겹의 보컬 화음입니다.
newage|사람의 노래 없이 피아노만 들리는 연주곡입니다.
`);
group('목소리','목소리의 높낮이',`
cash|낮고 깊게 울리는 목소리의 질감입니다.
darkvoice|아래 음역에 머무는 묵직한 보컬입니다.
soprano|매우 높은 음역의 선율을 들어보세요.
bassvoice|베이스 성악가의 낮은 음역과 공명입니다.
baritone|바리톤 성악가의 중저음과 표현입니다.
opera|테너 성악가의 밝고 높은 남성 음역입니다.
alto|콘트랄토의 낮고 풍부한 여성 음역을 참고하세요.
mezzo|메조소프라노의 중간 여성 음역과 질감입니다.
soprano|소프라노의 높은 음역과 성악적 공명입니다.
falsetto|가볍고 높은 남성 보컬의 가성 음색입니다.
`);
group('목소리','부드럽고 감성적인 목소리',`
soft|힘을 과하게 주지 않는 부드러운 목소리입니다.
gentle|온화하게 이어지는 보컬의 호흡입니다.
smooth|포근하고 둥근 목소리의 울림입니다.
sweet|달콤하고 가벼운 발음과 음색입니다.
easy|매끄럽게 이어지는 보컬의 흐름입니다.
balladkr|소리의 크기와 음색으로 감정을 쌓습니다.
love|다정하게 말을 건네는 듯한 표현입니다.
breathy|보컬에 섞이는 숨과 작은 소리를 들어보세요.
airy|가볍게 뜨는 공기감과 보컬의 여백입니다.
whisper|가까이서 속삭이는 듯한 보컬의 질감입니다.
`);
group('목소리','강렬하고 힘 있는 목소리',`
power|후렴의 큰 울림과 강한 지지감을 들어보세요.
aretha|단단하고 힘 있게 밀어내는 보컬입니다.
husky|조금 거칠고 탁한 결이 섞이는 음색입니다.
raspy|쉰 듯 까슬한 목소리의 질감입니다.
gritty|거친 소리와 강하게 뻗는 표현입니다.
raw|다듬지 않은 듯 생생하게 드러나는 보컬입니다.
whiskey|길게 꺾는 음과 깊은 감정의 표현입니다.
power|후렴에서 높고 힘 있게 뻗는 보컬을 참고하세요.
shout|외치듯 강하게 발음하는 보컬입니다.
scream|곡 후반의 거친 스크리밍 구간을 참고하세요.
`);
group('목소리','노래 부르는 방식(창법)',`
vibrato|길게 유지하는 음의 규칙적인 떨림입니다.
falsetto|가볍게 높은 음을 내는 보컬의 색깔입니다.
melisma|한 음절 안에서 여러 음으로 움직이는 장식입니다.
legato|음과 음 사이를 매끈하게 연결하는 성악 예시입니다.
staccato|짧게 끊어 부르는 빠른 성악 구절을 들어보세요.
whisper|낮은 강도와 숨 섞인 발성의 느낌입니다.
rap|음정보다 말의 리듬과 강세에 집중하세요.
spoken|선율보다 말의 억양으로 전달하는 표현입니다.
chant|반복되는 성가의 선율과 집단적인 울림입니다.
opera|넓게 울리는 공명과 성악적인 발성입니다.
`);
group('목소리','나이와 느낌을 표현하는 목소리',`
youth|가볍고 생기 있는 보컬 톤의 예시입니다.
adele|성숙하게 감정을 담아내는 목소리입니다.
child|어린 안나의 도입 보컬을 참고하세요. 실제 나이보다 표현의 느낌입니다.
cash|세월의 질감이 묻어나는 낮은 목소리입니다.
smooth|고전적인 가수의 발음과 보컬 질감입니다.
retrovoice|복고풍 댄스 음악에 어울리는 보컬 표현입니다.
clear|맑고 또렷하게 이어지는 발음을 들어보세요.
bright|밝게 뻗는 음색과 경쾌한 발음입니다.
darkvoice|낮고 어둡게 가라앉는 목소리입니다.
`);
group('악기','건반 악기',`
piano|피아노 한 대가 만드는 선율과 화음을 들어보세요.
grand|그랜드 피아노 독주의 넓은 음역과 울림입니다.
epiano|펑키한 전자 건반 음색을 들어보세요. 여러 전자 악기가 섞인 예시입니다.
keyboard|도입을 이끄는 선명한 키보드 리프입니다.
takeonme|반복되는 대표 멜로디의 전자 음색입니다.
organ|길게 유지되는 관악 같은 건반의 울림입니다.
hammond|해먼드 오르간의 두툼하고 부드러운 리프입니다.
accordion|아코디언 편곡으로 연주한 탱고 선율을 찾아보세요.
harpsichord|현을 뜯는 듯 짧고 또렷한 건반 소리입니다.
`);
group('악기','기타 및 현악기',`
guitar|통기타의 뜯는 소리와 몸통의 울림입니다.
electric|도입의 선명한 일렉트릭 기타 리프입니다.
classicalguitar|나일론 줄의 부드러운 소리와 트레몰로입니다.
bassguitar|일렉트릭 베이스의 독주와 하모닉스입니다.
bassguitar|전기 베이스의 저음과 특수한 배음을 들어보세요.
violin|높고 민첩하게 움직이는 바이올린 선율입니다.
viola|바이올린보다 낮고 두툼한 현악 음색입니다.
cello|낮고 따뜻하게 울리는 첼로 독주입니다.
doublebass|큰 현악기의 낮은 울림과 독주 선율입니다.
harp|줄을 뜯어 물결처럼 번지는 하프의 울림입니다.
ukulele|작고 가볍게 튕기는 우쿨렐레 반주입니다.
mandolin|짧고 또렷하게 튕기는 만돌린의 음색입니다.
banjo|밝고 금속성인 줄 소리와 빠른 반복입니다.
`);
group('악기','드럼 및 타악기',`
drums|드럼 솔로에서 여러 북과 심벌의 소리를 비교하세요.
drums|킥·스네어·탐·심벌이 함께 움직이는 드럼 세트입니다.
edrums|전자 타악기 질감과 후반의 어쿠스틱 드럼을 비교하세요.
kick|바닥을 규칙적으로 짚는 낮은 둥둥 소리입니다.
snare|작품 내내 이어지는 스네어의 짧고 선명한 반복입니다.
hihat|잘게 쪼개지는 하이햇과 셔플 그루브입니다.
cymbals|합주 속에서 밝게 퍼지는 심벌의 충돌음을 들어보세요.
tambourine|반주 위에 흔들리며 반짝이는 금속 타악 소리입니다.
bongo|작은 손북의 높고 통통 튀는 타격감입니다.
conga|손으로 두드리는 둥근 북소리와 라틴 그루브입니다.
timpani|도입부의 낮고 음정이 있는 북소리입니다.
xylophone|짧고 단단한 나무 건반의 타격음입니다.
marimba|실로폰보다 둥글고 풍부하게 울리는 나무 건반입니다.
`);
group('악기','관악기',`
flute|숨이 흐르는 맑고 유연한 플루트 독주입니다.
piccolo|합주 위로 높게 떠오르는 작은 플루트의 소리입니다.
clarinet|부드럽고 둥근 목관의 선율입니다.
oboe|약간 콧소리처럼 선명한 리드 악기의 음색입니다.
bassoon|낮고 두툼하며 익살스러운 목관의 음색입니다.
sax|도입의 굵고 노래하는 듯한 색소폰 선율입니다.
trumpet|앞으로 뻗는 금관의 선명한 소리입니다.
trombone|슬라이드 악기의 부드럽고 묵직한 선율입니다.
horn|둥글게 감싸는 금관의 넓은 울림입니다.
tuba|금관악기 중 낮고 큰 울림을 들어보세요.
harmonica|숨을 불고 들이쉬며 꺾는 작은 리드 악기 소리입니다.
recorder|부드럽고 맑은 리코더의 빠른 선율입니다.
panflute|대나무 관을 넘나드는 숨 섞인 선율입니다.
`);
group('악기','국악기',`
gayageum|줄을 뜯은 뒤 음을 흔드는 가야금의 농현입니다.
geomungo|술대로 줄을 치고 뜯는 묵직한 거문고 소리입니다.
haegeum|활로 켜는 두 줄의 가늘고 애절한 음색입니다.
daegeum|넓은 숨소리와 청의 떨림이 섞이는 대금입니다.
piri|작은 관에서 나오는 선명하고 힘 있는 소리입니다.
janggu|양쪽 북면의 서로 다른 높이와 장단입니다.
buk|소리꾼 옆에서 장단을 받치는 북 반주입니다.
kkwaenggwari|사물놀이를 이끄는 높고 날카로운 금속음입니다.
jing|낮고 길게 퍼지는 큰 금속 타악기의 울림입니다.
taepyeongso|야외로 멀리 뻗는 강하고 선명한 관악 소리입니다.
danso|작은 대나무 관의 맑고 고요한 음색입니다.
`);
group('악기','전자 음악 악기 및 사운드',`
synthbass|반복되는 낮은 신시사이저 베이스를 들어보세요.
takeonme|앞에 나와 주선율을 담당하는 전자음입니다.
synthpad|뒤에서 길게 지속하며 공간을 채우는 전자음입니다.
machine|기계적으로 반복되는 전자 타악기 패턴입니다.
bass808|낮고 길게 울리는 전자 베이스와 킥의 질감입니다.
house|전자적으로 반복되는 킥과 비트입니다.
arpeggio|화음의 음들이 차례로 반복되는 신스 패턴입니다.
stringmachine|현악처럼 길게 깔리는 전자 음색을 들어보세요.
`);
group('악기','오케스트라 악기 구성',`
strings|바이올린·비올라·첼로·베이스의 현악 합주입니다.
mournful|여러 현악 성부가 층을 이루는 울림입니다.
brass|트럼펫·호른·트롬본·튜바가 만드는 금관 합주입니다.
woodwind|목관 5중주 편곡으로 음색의 조합을 들어보세요.
dvorak|현악·목관·금관·타악 전체가 어우러지는 소리입니다.
strings|오케스트라 현악군의 넓고 풍성한 소리입니다.
orchestralperc|행진곡 속 타악기의 강한 박자와 강조입니다.
`);
group('노래 구조','기본적인 노래 구조',`
takeonme|보컬이 들어오기 전 신스 리프로 곡을 소개합니다.
adele|이야기를 전하는 절과 후렴의 선율을 비교하세요.
adele|처음 등장하는 절을 듣고 첫 후렴과 비교하세요.
adele|첫 후렴 뒤 돌아오는 절의 가사와 반주 변화를 들어보세요.
electro|후렴 직전에 기대감을 높이는 반복 구간입니다.
dynamite|곡에서 되돌아오는 핵심 선율을 찾아보세요.
postchorus|후렴 뒤 이어지는 반복 보컬 구간을 찾아보세요.
bohemian|곡의 새로운 분위기로 넘어가는 대비를 참고하세요. 전형적인 팝 브리지와는 다릅니다.
eagles|마지막 보컬 뒤 이어지는 긴 기타 연주입니다.
beethoven|1악장의 마지막 종지로 음악이 끝나는 느낌을 들어보세요.
`);
group('노래 구조','연주 구간을 지정하는 태그',`
newage|가사 없이 피아노만 이어지는 곡입니다.
electric|보컬 이전 기타가 곡의 첫인상을 만듭니다.
eagles|보컬이 쉬고 악기가 앞에 나오는 구간을 찾아보세요.
november|노래 사이와 후반의 기타 독주를 들어보세요.
grand|피아노 한 대가 곡을 이끄는 독주 예시입니다.
sax|보컬 사이에 되돌아오는 색소폰 선율입니다.
drums|합주 중 드럼이 주인공이 되는 솔로 구간입니다.
`);
group('노래 구조','노래의 특별한 구간',`
takeonme|도입부터 기억에 남는 신스 리프가 훅입니다.
folk|각 절 끝에 되돌아오는 짧은 선율을 들어보세요.
badguy|곡 후반의 리듬과 음색이 달라지는 전환을 참고하세요.
build|큰 비트가 나오기 전 긴장감이 쌓이는 흐름입니다.
drop|고조된 뒤 중심 리듬과 베이스가 터지는 부분입니다.
interlude|노래 흐름 사이에 등장하는 말과 효과음의 연결입니다.
power|마지막 후렴에서 높아지는 감정과 편곡입니다.
`);
group('노래 구조','반복 및 마무리 표현',`
fade|마지막에 되풀이되는 긴 합창 구절을 들어보세요.
eagles|끝부분에서 계속 이어지는 기타 후주입니다.
fade|곡 끝으로 갈수록 전체 음량이 작아집니다.
beethoven|마지막 화음과 멈춤이 만드는 명확한 끝맺음입니다.
`);
let category='', subgroup='';
const rows=[];
const positions=new Map();
for (const line of fs.readFileSync(path.join(dir,'suno_simple.txt'),'utf8').split(/\r?\n/)) {
  const l=line.trim();
  if (l.endsWith('{')) { category=l.slice(0,-1); subgroup=''; }
  else if(l.includes('#')) {
    const [tag,ko]=l.split(/\s*#\s*/);
    if(category==='장르' && tag.trim()==='Classical') subgroup='클래식 · 영화 · 무대';
    const key=category+'/'+subgroup, index=positions.get(key)||0;
    const example=mappings.get(key)?.[index];
    if(!example) throw new Error('Missing example: '+key+' '+tag);
    rows.push({id:rows.length,category,subgroup,tag:tag.trim(),ko:ko.trim(),...example});
    positions.set(key,index+1);
  } else if(l && l!=='}' && !l.startsWith('(')) subgroup=l;
}
for(const [key,items] of mappings) if(positions.get(key)!==items.length) throw new Error('Length mismatch '+key);
const categories=[...new Set(rows.map(r=>r.category))];
const payload={rows,songs,categories};
let html=fs.readFileSync(path.join(dir,'template.html'),'utf8');
html=html.replace('/*__DATA__*/',JSON.stringify(payload).replace(/</g,'\\u003c'));
fs.writeFileSync(path.join(dir,'index.html'),html,'utf8');
console.log(JSON.stringify({file:'index.html',entries:rows.length,uniqueTags:new Set(rows.map(r=>r.category+'/'+r.tag)).size,songs:Object.keys(songs).length,directVideos:Object.values(songs).filter(s=>s.video).length,bytes:Buffer.byteLength(html)},null,2));
