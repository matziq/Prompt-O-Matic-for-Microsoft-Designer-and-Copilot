<#
.SYNOPSIS
This script creates a prompt form for generating prompts for Microsoft Designer.

.DESCRIPTION
The script creates a Windows Forms application that allows the user to generate prompts for Microsoft Designer. It provides dropdowns for selecting the subject, action, and artistic style of the prompt.

.PARAMETER None

.INPUTS
None

.OUTPUTS
None

.EXAMPLE
.\Designer Prompt O Matic.ps1
This command runs the script and opens the prompt form.

.NOTES
Author: John Rea
Date: July 29, 2024
Version: 1.5
#>

# Adds the System.Windows.Forms assembly.
Add-Type -AssemblyName System.Windows.Forms

# Change the working directory to the directory where the script is located
Set-Location -Path $PSScriptRoot

# Verify the change by getting the current working directory
$currentDirectory = Get-Location
Write-Output "Current working directory: $currentDirectory"

# Clears the clipboard on launch
[System.Windows.Forms.Clipboard]::Clear()

function Format-ForURL ($text) {
    try {
        $text = $text.Replace(' ','+')
        return $text
    } catch {
        [System.Windows.Forms.MessageBox]::Show('An error occurred while formatting the text: $_','Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Prompt-O-Matic for Microsoft Designer'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(600, 550)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $true
$form.MinimizeBox = $true

# Look and Feel
$form.BackColor = [System.Drawing.Color]::LightSteelBlue # Background color
$form.ForeColor = [System.Drawing.Color]::DarkSlateGray # Text color

# Create dropdown for Subject
$subjectLabel = New-Object System.Windows.Forms.Label
$subjectLabel.Text = 'Subject'
$subjectLabel.AutoSize = $true
$subjectLabel.Location = New-Object System.Drawing.Point(5, 10)
$subjectCheckboxLabel = New-Object System.Windows.Forms.Label
$subjectCheckboxLabel.Location = New-Object System.Drawing.Point(130, 20)
$subjectCheckboxLabel.AutoSize = $true
$subjectCheckbox = New-Object System.Windows.Forms.CheckBox
$subjectCheckbox.Location = New-Object System.Drawing.Point(180, 30)
$subjectCheckbox.text = 'Freeze',$subjectLabel.Text
$subjectCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($subjectCheckbox)
$subjectDropdown = New-Object System.Windows.Forms.ComboBox
$subjectDropdown.Location = New-Object System.Drawing.Point(5, 30)
$subjectDropdown.Size = New-Object System.Drawing.Size(165, 20)
$subjectDropdown.DropDownStyle = 'DropDownList'
$subjectDropdown.AutoCompleteMode = 'None'
$subjectDropdown.AutoCompleteSource = 'ListItems'
$subjectdropdown.items.addrange((@('actor','airplane','alligator','animal','antelope','archipelago','backyard','balloon','banksys defy','bay','beach','bear','bedroom','bicycle','book','boy','bridge','building','butterfly','camera','cape','car','castle','cat','cave','child','city','cliff','cloud','computer','conflict','coral','crocodile','cultural symbol','da vincis gaze','dalis melt','dancer','den','desert','dessert','doctor','dolphin','dream','eagle','elephant','fantasy world','flower','forest','fountain','frog','galaxy','garden','garage','giraffe','girl','glacier','guitar','harbor','hero','hill','hippopotamus','historical event','hokusais wave','hoppers isolate','horse','iceberg','island','jellyfish','kangaroo','kitchen','kite','klimts embrace','koala','lamp','landscape','lighthouse','lion','living room','love','man','matisses cut','meadow','michelangelos creation','monkey','monets bloom','moon','mountain','museum','musician','mythological creature','nightmare','ocean','octopus','okeeffes enlarge','owl','panda','parrot','patio','peace','peacock','pebble','penguin','peninsula','piano','pirate','planet','pond','pollocks drip','portrait','pyramid','rabbit','rain','ranger','religious icon','rembrandts illuminate','river','sand','science fiction scene','snow','star','storm','street','tree','truck','turtle','tyrannosaurus rex','valley','van goghs swirl','villain','volcano','warhols replicate','waterfall','whale','wildabeest','willem de koonings woman','wolf','woman','zebra' ) | foreach-object {'A [' + $_ + ']'} | sort-object -unique))

# Create dropdown for Action
$actionLabel = New-Object System.Windows.Forms.Label
$actionLabel.Text = 'Action'
$actionLabel.AutoSize = $true
$actionLabel.Location = New-Object System.Drawing.Point(5, 60)
$actionCheckbox = New-Object System.Windows.Forms.CheckBox
$actionCheckbox.Location = New-Object System.Drawing.Point(180, 80)
$actionCheckbox.Text = 'Freeze',$actionLabel.Text
$actionCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($actionCheckbox)
$actionDropdown = New-Object System.Windows.Forms.ComboBox
$actionDropdown.Location = New-Object System.Drawing.Point(5, 80)
$actionDropdown.DropDownStyle = 'DropDownList'
$actionDropdown.AutoCompleteMode = 'None'
$actionDropdown.AutoCompleteSource = 'ListItems'
$actionDropdown.Size = New-Object System.Drawing.Size(165, 20)
$actiondropdown.items.addrange((@('admiring','aiding','allowing','analyzing','appreciating','appointing','arguing','arranging','aspiring','assembling','assisting','auditing','authorizing','awakening','backing','bargaining','bartering','becoming','building','calculating','campaigning','capturing','caring','carving','celebrating','certifying','changing','charging','cherishing','choosing','climbing','coding','commanding','communicating','compensating','competing','completing','concluding','confirming','connecting','considering','constructing','contemplating','controlling','cooking','coordinating','creating','curing','cutting','cycling','dancing','debating','deciding','defending','defining','delivering','designing','detecting','developing','directing','discussing','discovering','distributing','diving','dock','drawing','dreaming','driving','editing','educating','electing','emancipating','employing','empowering','enabling','endorsing','energizing','enjoying','equipping','etching','examining','existing','explaining','exploring','expressing','modeting','feeling','filming','financing','fishing','flying','forging','forgetting','functioning','funding','governing','greeting','growing','guarding','guiding','hacking','harvesting','healing','hearing','helping','highlighting','hiring','holding','honoring','hoping','hunting','identifying','illuminating','imagining','impacting','informing','innovating','inspecting','interpreting','investigating','jumping','keeping','labeling','landing','leading','leasing','learning','listening','living','loaning','lobbying','loving','maintaining','managing','manufacturing','marketing','measuring','mediating','medicating','mining','monitoring','molding','naming','navigating','negotiating','observing','offering','operating','orbiting','organizing','painting','performing','permitting','picking','planning','planting','playing','practicing','preparing','printing','processing','producing','programming','promoting','protecting','providing','publishing','racing','raising','reading','recalling','recording','recovering','refining','remembering','renting','rescuing','researching','resolving','respecting','rewarding','revitalizing','ruling','running','sailing','saving','sculpting','securing','seeing','selecting','selling','sensing','serving','shaping','shielding','shipping','singing','skating','skiing','smelling','solving','sponsoring','steering','stirring','studying','supporting','surfing','swimming','teaching','tracking','trading','trapping','traveling','treating','underlining','upholding','validating','viewing','voting','watching','winning','wishing','working','writing' ) | foreach-object {' that is [' + $_ + ']'} | sort-object -unique))

# Create dropdown for Artistic Style
$styleLabel = New-Object System.Windows.Forms.Label
$styleLabel.Text = 'Artistic Style'
$styleLabel.AutoSize = $true
$styleLabel.Location = New-Object System.Drawing.Point(5, 110)
$styleCheckbox = New-Object System.Windows.Forms.CheckBox
$styleCheckbox.Location = New-Object System.Drawing.Point(180, 130)
$styleCheckbox.Text = 'Freeze',$styleLabel.Text
$styleCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($styleCheckbox)
$styleDropdown = New-Object System.Windows.Forms.ComboBox
$styleDropdown.Location = New-Object System.Drawing.Point(5, 130)
$styleDropdown.DropDownStyle = 'DropDownList'
$styleDropdown.AutoCompleteMode = 'None'
$styleDropdown.AutoCompleteSource = 'ListItems'
$styleDropdown.Size = New-Object System.Drawing.Size(165, 20)
$styledropdown.items.addrange((@('abstract art','abstract expressionism','academic art','acrylic paint','algorithmic art','antique','arte povera','art brut','art deco','art nouveau','augmented reality art','avant-garde','baroque','bauhaus','bio art','biological materials','bronze','canvas','caravaggism','ceramics','charcoal','classicism','clay','cobr','collage','color field painting','colored pencil','conceptual art','constructivism','contemporary art','cubism','cyber','cyber art','dadaism','de stijl','digital art','encaustic','environmental art','environmental materials','ephemeral materials','expressionism','fauvism','fiber','figurative art','fine art','found objects','futurism','futuristic','generative art','glass','gothic','graphite','graffiti','gouache paint','harlem renaissance','impressionism','ink','installation art','land art','leather','light','light art','lowbrow','lyrical abstraction','marble','medieval','metal','minimalism','modern','modern art','modernism','mosaic','naive art','nautical','neo-dada','neo-expressionism','neo-geo','neo-impressionism','neoclassicism','neon art','oil paint','op art','paper','pastel','performance art','performance elements','photorealism','picassian','plastics','pop art','porcelain','post-impressionism','post-internet','post-structuralism','postmodernism','precisionism','realism','renaissance','rembrandt lighting','relational aesthetics','retro','rococo','romantic','romanticism','sand','silverpoint','sound','sound art','steampunk','stone','street art','stuckism','superflat','suprematism','surrealism','symbolism','tachisme','tempera paint','textiles','toyism','transavantgarde','transavanguardia','video','video art','vintage','virtual reality art','watercolor paint','wood','young british artists','zero group' ) | foreach-object {' in the style of [' + $_ + ']'} | sort-object -unique))

# Create dropdown for Artistic Media
$mediaLabel = New-Object System.Windows.Forms.Label
$mediaLabel.Text = 'Artistic Media'
$mediaLabel.AutoSize = $true
$mediaLabel.Location = New-Object System.Drawing.Point(5, 160)
$mediaCheckbox = New-Object System.Windows.Forms.CheckBox
$mediaCheckbox.Location = New-Object System.Drawing.Point(180, 180)
$mediaCheckbox.Text = 'Freeze',$mediaLabel.Text
$mediaCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($mediaCheckbox)
$mediaDropdown = New-Object System.Windows.Forms.ComboBox
$mediaDropdown.Location = New-Object System.Drawing.Point(5, 180)
$mediaDropdown.DropDownStyle = 'DropDownList'
$mediaDropdown.AutoCompleteMode = 'None'
$mediaDropdown.AutoCompleteSource = 'ListItems'
$mediaDropdown.Size = New-Object System.Drawing.Size(165, 20)
$mediadropdown.items.addrange((@('acrylic paint','acrylic resin','airbrush','alabaster','bamboo','batik','beads','beeswax','biologicals','bone','brass','bronze','buttons','calligraphy','car parts','cement','ceramic','chalk','charcoal','clay','collage','colored pencils','concrete','copper','coral','crayon','dammar','diamonds','digital art','dirt','dye','egg tempera','embroidery','encaustic','environmental materials','ephemeral materials','etching','fabric','feathers','felt','fiber','film','fimo','flint','floral foam','foam','foil','found objects','fresco','fur','garbage','gesso','glass','glitter','glue','gold','gold leaf','gouache paint','granite','graphite','graphite powder','gypsum','hair','hemp','ice','ink','iron','ivory','jade','jewelry','kinetic art','kite','knitting','lacquer','lace','latex','lead','leather','leaves','light','limestone','linoleum','lino','lithography','mache','magnet','mahogany','marble','marbling','markers','masks','matchsticks','metal','mirror','mixed media','modeling clay','moss','mosaic','mylar','nails','neon','oak','obsidian','oil paint','origami','papyrus','paper','paper mache','paraffin','pastel','pastel pencils','pearls','pen','pencil','pewter','photography','plaster','plaster of paris','plastics','plexiglass','polymer','polymer clay','porcelain','porphyry','pottery','printmaking','quilling','quilt','recycled materials','resin','ribbon','rocks','rubber','salt','sand','sandpaper','sawdust','scissors','scrapbooking','sculpture','seashells','sewing','shells','shrinking plastic','silk','silver','silverpoint','slate','soap','soapstone','sound','stained glass','stamps','stencils','steel','stickers','stone','straw','string','tape','tar','tempera paint','textiles','terracotta','thread','tissue paper','utensils','vellum','velvet','venetian plaster','video','vinyl','wax','watercolor','wire','wood','wool','yarn','zinc' ) | foreach-object {' that is using [' + $_ + ']'} | sort-object -unique))

# Create dropdown for Artist Color
$colorLabel = New-Object System.Windows.Forms.Label
$colorLabel.Text = 'Color Palette'
$colorLabel.AutoSize = $true
$colorLabel.Location = New-Object System.Drawing.Point(5, 210)
$colorCheckbox = New-Object System.Windows.Forms.CheckBox
$colorCheckbox.Location = New-Object System.Drawing.Point(180, 230)
$colorCheckbox.Text = 'Freeze',$colorLabel.Text
$colorCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($colorCheckbox)
$colorDropdown = New-Object System.Windows.Forms.ComboBox
$colorDropdown.Location = New-Object System.Drawing.Point(5, 230)
$colorDropdown.DropDownStyle = 'DropDownList'
$colorDropdown.AutoCompleteMode = 'None'
$colorDropdown.AutoCompleteSource = 'ListItems'
$colorDropdown.Size = New-Object System.Drawing.Size(165, 20)
$colordropdown.items.addrange((@('amber','analogous','arctic','autumn','azure','beach','berry','black and white','blossom','blue','blush','bold','bronze','candy','carnival','cerulean','charcoal','cherry','chocolate','cinnamon','citrus','cobalt','coffee','complementary','cool','copper','coral','cream','crimson','crystal','dark','desaturated','desert','dull','earth tones','emerald','fire','floral','forest','frost','fuchsia','galactic','glow','gold','grayscale','green','honey','ice','indigo','ivory','jade','jewel tones','khaki','lagoon','lavender','lemon','lilac','lime','light','magenta','mango','marble','maroon','metallic','merlot','midnight','mint','monochrome','moss','muted','nature','navy','neon','neutral','oatmeal','oceanic','ochre','olive','olivine','onyx','opal','orchid','pastel','papaya','peach','pearl','periwinkle','pink','pine','plum','primary','psychedelic','pumpkin','purple','rainbow','red','retro','romantic','rose','rose gold','royal','ruby','rustic','saffron','sage','salmon','sand','sapphire','saturated','scarlet','seashell','secondary','sepia','sienna','silver','sky','slate','smoke','smoky','snow','spicy','split-complementary','spring','stormy','strawberry','summer','sunrise','sunset','tangerine','taupe','teal','tertiary','tetradic','thistle','topaz','triadic','tropical','turquoise','twilight','urban','vanilla','velvet','vermilion','vibrant','violet','warm','watermelon','wheat','whimsical','wine','winter','yellow','zinc' ) | foreach-object {' with a color palette of [' + $_ + ']'} | sort-object -unique))

# Create dropdown for Artist Mood
$moodLabel = New-Object System.Windows.Forms.Label
$moodLabel.Text = 'Artist Mood'
$moodLabel.AutoSize = $true
$moodLabel.Location = New-Object System.Drawing.Point(5, 260)
$moodCheckbox = New-Object System.Windows.Forms.CheckBox
$moodCheckbox.Location = New-Object System.Drawing.Point(180, 280)
$moodCheckbox.Text = 'Freeze',$moodLabel.Text
$moodCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($moodCheckbox)
$moodDropdown = New-Object System.Windows.Forms.ComboBox
$moodDropdown.Location = New-Object System.Drawing.Point(5, 280)
$moodDropdown.DropDownStyle = 'DropDownList'
$moodDropdown.AutoCompleteMode = 'None'
$moodDropdown.AutoCompleteSource = 'ListItems'
$moodDropdown.Size = New-Object System.Drawing.Size(165, 20)
$mooddropdown.items.addrange((@('8k','abstract','adventurous','aerial','analog','astrophotography','backlit','black and white','bohemian','bokeh','burst mode','candid','calming','cinematic','collage','colorful','complementary colors','cozy','creative','cropped','crisp','curves','dark','depth of field','diptych','distorted','double exposure','dramatic','dreamlike','dreamy','duotone','dynamic','edgy','elegant','emotional','enchanting','ethereal','exotic','expressive','extreme close-up','extremely detailed','fantasy','film noir','fish-eye','flashy','framed','frozen','futuristic','geometric','glamorous','glittery','glossy finish','glowing','gothic','golden hour','grainy','gritty','grunge','hazy','high contrast','high dynamic range','high key','horror','humorous','illuminated','impressionist','incandescent','infrared','intense','intricate','isolated','joyful','kodachrome','landscape','leading lines','lens flare','lively','long exposure','low key','low-light','luminous','macro','magical','majestic','matte finish','metallic','minimalist','monochrome','monumental','moody','motion blur','mysterious','natural','neon','night','nostalgic','oil painting','old-fashioned','optical illusion','ornate','painterly','panoramic view','pastel','patterned','perspective','photorealistic','playful','polaroid','pop art','portrait','quirky','rainbow','realistic','reflections','retro','rich','romantic','rule of thirds','rustic','saturation','scenic','sepia','serene','shadow','shallow depth of field','sharp','shiny','silhouette','smoke filled','smooth','soft focus','solarized','sparkling','split tone','spooky','star trails','still life','storytelling','studio lighting','stunning','subtle','sunrise','sunset','surreal','symmetry','textured','tilt-shifted','time-lapse','tinted','tonal','tranquil','trippy','underwater','upside down','urban','vibrant','vignette','vintage','vivid','warm','watercolor','wet','whimsical','wild','wonderful','zoomed' ) | foreach-object {' and an artistic mood of [' + $_ + ']'} | sort-object -unique))

# Create dropdown for Artist Feeling
$feelingLabel = New-Object System.Windows.Forms.Label
$feelingLabel.Text = 'Artist Feeling'
$feelingLabel.AutoSize = $true
$feelingLabel.Location = New-Object System.Drawing.Point(5, 310)
$feelingCheckbox = New-Object System.Windows.Forms.CheckBox
$feelingCheckbox.Location = New-Object System.Drawing.Point(180, 330)
$feelingCheckbox.Text = 'Freeze',$feelingLabel.Text
$feelingCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($feelingCheckbox)
$feelingDropdown = New-Object System.Windows.Forms.ComboBox
$feelingDropdown.Location = New-Object System.Drawing.Point(5, 330)
$feelingDropdown.DropDownStyle = 'DropDownList'
$feelingDropdown.AutoCompleteMode = 'None'
$feelingDropdown.AutoCompleteSource = 'ListItems'
$feelingDropdown.Size = New-Object System.Drawing.Size(165, 20)
$feelingDropdown.items.addrange((@('3d rendered','aerial view','ambient light','backlit','birds eye view','bokehlicious','chiaroscuro','cross processing','cyberpunk','deep depth of field','dystopian','fantasy','film grain','fish-eye lens','flat lay','hard shadows','isometric','lens flare','light painting','macro shot','magical realism','negative space','panoramic','perspective','post-apocalyptic','rim light','sci-fi','silhouette','soft shadows','split toning','steampunk','sunburst','telephoto','utopian','wide angle','worms eye view','zoomed in','abstract','allegorical','aperture','blue color temperature','blue hour','brightness','candlelight','conceptual','contrast','dawn','depressive tone','diurnal','dusk','figurative','futuristic','geometric','glow','golden hour','high aperture','high iso','high shutter speed','historical','illumination','long focal length','low exposure','luminance','medium saturation','medium white balance','moonlight','mythological','narrative','neon','nightfall','nocturnal','opposite hue','organic','out of focus','radiance','retro','seasonal','starlight','symbolic','timeless','twilight','analog','angular','antique','backdrop','bokeh','burst','candid','cinemagraph','collage','colorful','complementary','composition','cropped','curves','depth','distortion','double exposure','dramatic','duotone','exposure','fade','filter','flash','focus','framing','glare','gradient','grayscale','HDR','horizon','hue','infrared','low light','macro','monochrome','negative','noise','overexposed','polaroid','refraction','resolution','saturation','sepia','shadows','sharpness','snapshot','strobe','tilt-shift','tone','underexposed','vibrance','vignette','vintage','warm','watermark','zoom','backlighting','bloom','blur','boomerang','brenizer','calotype','clarity','collodion','cross polarization','daguerreotype','darkroom','digital','dodge and burn','dreamy','eerie','ethereal','faded','ferrotype','flare','focal point','frame within a frame','ghosting','grainy','halo','haze','high dynamic range','high key','hyperlapse','hypnotic','impasto','inverted','kinetic','lensbaby','light leak','light trail','lithograph','long exposure','low key','luminous','matte','metallic','minimalist','mirror','moire','motion blur','negative fill','nostalgic','oil painting','old-fashioned','optical illusion','orb','overlapping','painterly','parallax','pastel','pinhole','polarization','pop art','prism','psychedelic','rainbow','reflection','rembrandt','retouch','romantic','rotoscope','saturated','selective color','shallow depth of field','sketch','slow shutter speed','solarization','sparkle','spooky','starburst','stencil','stitching','stop motion','surreal','texture','tilt','time-lapse','tint','trippy','vanishing point','vaporwave','vector','watercolor','wide aperture','zenith','zoom burst' ) | foreach-object {' and an artistic feeling of [' + $_ + ']'} | sort-object -unique))

# Create dropdown for Art Look & Feel
$modeLabel = New-Object System.Windows.Forms.Label
$modeLabel.Text = 'Art Look & Feel'
$modeLabel.AutoSize = $true
$modeLabel.Location = New-Object System.Drawing.Point(5, 360)
$modeCheckbox = New-Object System.Windows.Forms.CheckBox
$modeCheckbox.Location = New-Object System.Drawing.Point(180, 380)
$modeCheckbox.Text = 'Freeze',$modeLabel.Text
$modeCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($modeCheckbox)
$modeDropdown = New-Object System.Windows.Forms.ComboBox
$modeDropdown.Location = New-Object System.Drawing.Point(5, 380)
$modeDropdown.DropDownStyle = 'DropDownList'
$modeDropdown.AutoCompleteMode = 'None'
$modeDropdown.AutoCompleteSource = 'ListItems'
$modeDropdown.Size = New-Object System.Drawing.Size(165, 20)
$modedropdown.items.addrange((@('abstract photography','aperture priority mode','astrophotography','autofocus','backlighting','black and white photography','bokeh','bottomlighting','bracketing','bracketing mode','burst mode','candid photography','center-weighted metering','cold shoe','color photography','compact camera','continuous shot mode','dark filter','depth of field','digital camera','diffuser','dslr','evaluative metering','exposure compensation','exposure lock','exposure value','film camera','fish-eye lens','flash compensation','flash photography','focus','focus lock','food photography','found footage','frontlighting','f-stop','hard light','hdr','high aperture','high depth of field','high exposure','high focal length','histogram','hot shoe','hyperfocal distance','image stabilization','infrared photography','instamatic','iso','jpeg format','kodachrome','landscape photography','lens mount','light filter','light meter','lomography','long exposure','low aperture','low exposure','low focal length','macro lens','macro photography','manual focus','manual mode','matrix metering','medium aperture','medium exposure','medium filter','medium focal length','medium format camera','metering mode','mirrorless camera','monopod','narrow lens','neutral density filter','night photography','panorama','pinhole camera','pixel','point-and-shoot camera','polarizing filter','portrait photography','prime lens','program mode','rangefinder camera','raw format','reflector','remote trigger','resolution','self-timer','sensor','sepia photography','shaky cam','shutter priority mode','shutter speed','sidelighting','single shot mode','slr','soft light','softbox','speedlight','spot metering','sports photography','street photography','sync cord','telephoto lens','time-lapse','tlr','toplighting','tripod','viewfinder camera','white balance','wide-angle lens','wildlife photography','zoom lens' ) | foreach-object {' and an artistic mode of [' + $_ + '].'} | sort-object -unique))

# Add event handlers to dropdowns
$subjectDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$actionDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$styleDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$mediaDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$colorDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$moodDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$feelingDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$modeDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)

$dropdowns = @($subjectDropdown, $actionDropdown, $styleDropdown, $mediaDropdown, $colorDropdown, $moodDropdown, $feelingDropdown, $modeDropdown)

foreach ($dropdown in $dropdowns) {
    $dropdown.Add_SelectedIndexChanged({
        $selectedWords = @()
        foreach ($dd in $dropdowns) {
            if ($dd.SelectedItem) { $selectedWords += $dd.SelectedItem }
        }
        $clipboardText = [string]::Join(',', $selectedWords)
        $copiedTextbox.Text = $clipboardText
        {if ($clipboardText -ne $null) {
            Set-Clipboard -Value $clipboardText
        }}
    })
}
    $dropdown.Add_SelectedIndexChanged({
        $selectedWords = @()
        foreach ($dd in $dropdowns) {
            if ($dd.SelectedItem) { $selectedWords += $dd.SelectedItem }
        }
        $clipboardText = [string]::Join(',', $selectedWords)
        $copiedTextbox.Text = $clipboardText
        {if ($clipboardText -ne $null) {
            Set-Clipboard -Value $clipboardText
        }}
    })
foreach ($dropdown in $dropdowns) {
    $dropdown.Add_TextChanged({
        $selectedWords = @()
        foreach ($dd in $dropdowns) {
            if ($dd.Text) { $selectedWords += $dd.Text }
        }
        $clipboardText = [string]::Join(',', $selectedWords)
        $copiedTextbox.Text = $clipboardText
        {if ($clipboardText -ne $null) {
            Set-Clipboard -Value $clipboardText
        }}
    })
}

# Create label and textbox to display copied text
$copiedLabel = New-Object System.Windows.Forms.Label
$copiedLabel.Text = 'Current Prompt in the Clipboard'
$copiedLabel.AutoSize = $true
$copiedLabel.Location = New-Object System.Drawing.Point(10, 410)
$copiedTextbox = New-Object System.Windows.Forms.TextBox
$copiedTextbox.Location = New-Object System.Drawing.Point(10, 430)
$copiedTextbox.ReadOnly = $true
$copiedTextbox.Size = New-Object System.Drawing.Size(565, 60)
$copiedTextbox.Multiline = $true
$copiedTextbox.ScrollBars = 'Vertical'

# Create a 'Random' button
$randomButton = New-Object System.Windows.Forms.Button
$randomButton.Text = 'Random Prompt Generator'
$randomButton.Location = New-Object System.Drawing.Point(245, 30)
$randomButton.Size = New-Object System.Drawing.Size(150, 180)
$randomButton.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show('Are you sure? This will overwrite any chosen dropdowns that are not frozen!','Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmation -eq 'Yes') {
        $selectedWords = @()
        if(!$subjectCheckbox.Checked){
            $randomWord = $subjectDropdown.Items | Get-Random
            $selectedWords += $randomWord
            $subjectDropdown.SelectedItem = $randomWord
        }
        if(!$actionCheckbox.Checked){
            $randomWord = $actionDropdown.Items | Get-Random
            $selectedWords += $randomWord
            $actionDropdown.SelectedItem = $randomWord
        }
        if(!$styleCheckbox.Checked){
            $randomWord = $styleDropdown.Items | Get-Random
            $selectedWords += $randomWord
            $styleDropdown.SelectedItem = $randomWord
        }
        if(!$mediaCheckbox.Checked){
            $randomWord = $mediaDropdown.Items | Get-Random
            $selectedWords += $randomWord
            $mediaDropdown.SelectedItem = $randomWord
        }
        if(!$colorCheckbox.Checked){
            $randomWord = $colorDropdown.Items | Get-Random
            $selectedWords += $randomWord
            $colorDropdown.SelectedItem = $randomWord
        }
        if(!$moodCheckbox.Checked){
            $randomWord = $moodDropdown.Items | Get-Random
            $selectedWords += $randomWord
            $moodDropdown.SelectedItem = $randomWord
        }
        if(!$feelingCheckbox.Checked){
            $randomWord = $feelingDropdown.Items | Get-Random
            $selectedWords += $randomWord
            $feelingDropdown.SelectedItem = $randomWord
        }
        if(!$modeCheckbox.Checked){
            $randomWord = $modeDropdown.Items | Get-Random
            $selectedWords += $randomWord
            $modeDropdown.SelectedItem = $randomWord
        }
        $selectedWords = @()
        foreach ($dd in $dropdowns) {
            if ($dd.SelectedItem) { $selectedWords += $dd.SelectedItem }
        }
        $clipboardText = [string]::Join(',', $selectedWords)
        {if ($clipboardText -ne $null) {
            Set-Clipboard -Value $clipboardText
            }}
        {if ($clipboardText) {
            $copiedTextbox.Text = $clipboardText
        }}
    }
})

# Set the copied text to the joined selected words
if ($null -ne $selectedWords) {
    $copiedTextbox.Text = $selectedWords -join ' '
} else {
    if ($null -ne $copiedTextbox) {
        $copiedTextbox.Text = ''
    }
}

# Add a button to open the website with the generated text
$webMDButton = New-Object System.Windows.Forms.Button
$webMDButton.Text = 'Open Prompt in Microsoft Designer'
$webMDButton.Location = New-Object System.Drawing.Point(245, 215)
$webMDButton.Size = New-Object System.Drawing.Size(150, 90)

$webMDButton.Add_Click({
    $initialText = $copiedTextbox.Text
    if ($initialText -ne $null -and $initialText -ne '') {
        # Display an input box to allow text modification
        Add-Type -AssemblyName Microsoft.VisualBasic
        $modifiedText = [Microsoft.VisualBasic.Interaction]::InputBox("Modify the text if needed:", "Modify Text", $initialText)
        # Check if the user clicked "OK" without entering text
        if ($modifiedText -eq '') {
            [System.Windows.Forms.MessageBox]::Show('No text was entered. Please enter some text.','Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }
        else {
            [System.Windows.Forms.MessageBox]::Show('The clipboard is empty. Please select at least one item.','Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $baseURL = 'https://designer.microsoft.com/image-creator?p='
        # Assume Format-ForURL is a function that properly formats the text for a URL
        $formattedText = Format-ForURL -text $modifiedText

# Log the clipboard text to a file
$logFile = 'prompt.log'
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$logEntry = "$timestamp - $formattedText"
Add-Content -Path $logFile -Value $logEntry

if ($formattedText) {
    $edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    $profiles = @("Profile 8", "Profile 7", "Profile 6", "Profile 5", "Profile 2")
    foreach ($profile in $profiles) {
        $profileArgument = "--profile-directory=`"$profile`""
        $url = $baseURL + $formattedText + '&form=NTPCH1&refig=dd8bc5ce95bf495a89a1b6c447914a00&pc=U531&adppc=EDGEDBB&sp=2&lq=0&qs=PN&sk=PN1&sc=8-0&cvid=dd8bc5ce95bf495a89a1b6c447914a00&showconv=1&sendquery=1'
            # Validate both $profileArgument and $url before using them
            if (-not [string]::IsNullOrWhiteSpace($profileArgument) -and -not [string]::IsNullOrWhiteSpace($url)) {
                Start-Process -FilePath $edgePath -ArgumentList $profileArgument, $url
            } else {
                Write-Warning "Either profile argument or URL is null or empty."
            }
    }
}
})

# Add a button to open the website with the generated text
$webMCButton = New-Object System.Windows.Forms.Button
$webMCButton.Text = 'Open Prompt in Microsoft Copilot'
$webMCButton.Location = New-Object System.Drawing.Point(245, 310)
$webMCButton.Size = New-Object System.Drawing.Size(150, 90)
$webMCButton.Add_Click({
    if ($copiedTextbox.Text -ne $null -and $copiedTextbox.Text -ne '') {
        $clipboardText = Get-Clipboard
    } 
    else {
        [System.Windows.Forms.MessageBox]::Show('The clipboard is empty. Please select at least one item.','Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $baseURL = 'https://www.bing.com/search?q='
    # Open the website with the clipboard text
    $formattedText = Format-ForURL -text $copiedTextbox.Text
    if ($formattedText) {
        $url = $baseURL + 'an image of ' + $formattedText + '&form=NTPCH1&refig=dd8bc5ce95bf495a89a1b6c447914a00&pc=U531&adppc=EDGEDBB&sp=2&lq=0&qs=PN&sk=PN1&sc=8-0&cvid=dd8bc5ce95bf495a89a1b6c447914a00&showconv=1&sendquery=1&dissrchswrite=1'
        Start-Process $url
    } else {
        [System.Windows.Forms.MessageBox]::Show('The clipboard is empty. Please select at least one item.','Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }

# Log the clipboard text to a file
$logFile = 'prompt.log'
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$logEntry = "$timestamp - $formattedText"
Add-Content -Path $logFile -Value $logEntry
})

# Create a button to reset the dropdown boxes
$resetButton = New-Object System.Windows.Forms.Button
$resetButton.Location = New-Object System.Drawing.Point(400, 30)
$resetButton.Size = New-Object System.Drawing.Size(175, 370)
$resetButton.Text = 'Reset Dropdowns and Clear Clipboard'

# Add click event to the reset dropdown boxes button
$resetButton.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show('Are you sure? This will reset EVERYTHING!!','Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmation -eq 'Yes') {
        if ($subjectDropdown.Items.Count -gt 0) {$subjectDropdown.SelectedIndex = -1}
        if ($actionDropdown.Items.Count -gt 0) {$actionDropdown.SelectedIndex = -1}
        if ($styleDropdown.Items.Count -gt 0) {$styleDropdown.SelectedIndex = -1}
        if ($mediaDropdown.Items.Count -gt 0) {$mediaDropdown.SelectedIndex = -1}
        if ($colorDropdown.Items.Count -gt 0) {$colorDropdown.SelectedIndex = -1}
        if ($moodDropdown.Items.Count -gt 0) {$moodDropdown.SelectedIndex = -1}
        if ($feelingDropdown.Items.Count -gt 0) {$feelingDropdown.SelectedIndex = -1}
        if ($modeDropdown.Items.Count -gt 0) {$modeDropdown.SelectedIndex = -1}

        # Reset the checkboxes
        $subjectCheckbox.Checked = $false
        $actionCheckbox.Checked = $false
        $styleCheckbox.Checked = $false
        $mediaCheckbox.Checked = $false
        $colorCheckbox.Checked = $false
        $moodCheckbox.Checked = $false
        $feelingCheckbox.Checked = $false
        $modeCheckbox.Checked = $false
        if ([System.Windows.Forms.Clipboard]::ContainsText()) {
            [System.Windows.Forms.Clipboard]::Clear()
            $copiedTextbox.Text = ''
            $clearButton.Enabled = $false
    }}
})

# Add controls to the form
$form.Controls.Add($webMDButton)
$form.Controls.Add($webMCButton)
$form.Controls.Add($randomButton)
$form.Controls.Add($copiedLabel)
$form.Controls.Add($clearButton)
$form.Controls.Add($copybutton)
$form.Controls.Add($resetButton)
$form.Controls.AddRange(@($subjectLabel, $subjectDropdown, $subjectLabelText, $actionLabel, $actionDropdown, $styleLabel, $styleDropdown, $mediaLabel, $mediaDropdown, $colorLabel, $colorDropdown, $moodLabel, $moodDropdown, $feelingLabel, $feelingDropdown, $modeLabel, $modeDropdown, $copybutton, $copiedLabel, $copiedTextbox))
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

# Add event handler to prompt for confirmation when closing the form
$form.Add_FormClosing({
    param($sender, $eventArgs)
    $confirmation = [System.Windows.Forms.MessageBox]::Show('Are you sure you want to close the form?','Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmation -eq 'No') {
        $eventArgs.Cancel = $true
    } else {
        $eventArgs.Cancel = $false
    }
})

# Show form
$form.ShowDialog()
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'