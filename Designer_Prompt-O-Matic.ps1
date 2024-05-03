# DISCLAIMER
# This software (or sample code) is not supported under any Microsoft standard
# support program or service. The software is provided AS IS without warranty
# of any kind. Microsoft further disclaims all implied warranties including,
# without limitation, any implied warranties of merchantability or of fitness
# for a particular purpose. The entire risk arising out of the use or
# performance of the software and documentation remains with you. In no event
# shall Microsoft, its authors, or anyone else involved in the creation,
# production, or delivery of the software be liable for any damages whatsoever
# (including, without limitation, damages for loss of business profits, business
# interruption, loss of business information, or other pecuniary loss) arising
# out of the use of or inability to use the software or documentation, even if
# Microsoft has been advised of the possibility of such damages.
#
# Author: John Rea
# Date: May 1, 2024
# Version: 1.5

# Adds the System.Windows.Forms assembly.
Add-Type -AssemblyName System.Windows.Forms
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
$form.Topmost = $true

# Look and Feel
$form.BackColor = [System.Drawing.Color]::LightSteelBlue # Background color
$form.ForeColor = [System.Drawing.Color]::DarkSlateGray # Text color
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('C:\Windows\System32\notepad.exe') # Icon

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
$subjectDropdown.DropDownStyle = 'DropDown'
$subjectDropdown.Size = New-Object System.Drawing.Size(165, 20)
$subjectdropdown.items.addrange((@('abstract', 'actor', 'airplane', 'alligator', 'archipelago', 'eagle', 'elephant', 'iceberg', 'island', 'ocean', 'octopus', 'owl', 'animal', 'backyard', 'balloon', 'banksys defy', 'bay', 'beach', 'bear', 'bedroom', 'bicycle', 'book', 'boy', 'bridge', 'building', 'butterfly', 'camera', 'cape', 'car', 'castle', 'cat', 'cave', 'child', 'city', 'cliff', 'cloud', 'computer', 'conflict', 'coral', 'crocodile', 'cultural symbol', 'dancer', 'da vincis gaze', 'dalis melt', 'dancer', 'den', 'desert', 'dessert', 'doctor', 'dolphin', 'dream', 'fantasy world', 'flower', 'forest', 'fountain', 'frog', 'galaxy', 'garden', 'garage', 'giraffe', 'girl', 'glacier', 'guitar', 'harbor', 'hero', 'historical event', 'hill', 'hippopotamus', 'hokusais wave', 'hoppers isolate', 'horse', 'jellyfish', 'kangaroo', 'kitchen', 'kite', 'klimts embrace', 'koala', 'lamp', 'landscape', 'lighthouse', 'lion', 'living room', 'love', 'man', 'matisses cut', 'meadow', 'michelangelos creation', 'monkey', 'monets bloom', 'moon', 'mountain', 'museum', 'musician', 'mythological creature', 'nightmare', 'okeeffes enlarge', 'panda', 'parrot', 'patio', 'peace', 'peacock', 'pebble', 'penguin', 'peninsula', 'piano', 'pirate', 'planet', 'pond', 'pollocks drip', 'portrait', 'pyramid', 'rabbit', 'rain', 'ranger', 'religious icon', 'rembrandts illuminate', 'river', 'sand', 'science fiction scene', 'snow', 'star', 'storm', 'street', 'tree', 'truck', 'turtle', 'valley', 'van goghs swirl', 'villain', 'volcano', 'warhols replicate', 'waterfall', 'whale', 'wolf', 'woman', 'zebra') | sort-object -unique))

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
$actionDropdown.DropDownStyle = 'DropDown'
$actionDropdown.Size = New-Object System.Drawing.Size(165, 20)
$actiondropdown.items.addrange((@('admiring','aiding','allowing','analyzing','appreciating','appointing','arguing','arranging','aspiring','assembling','assisting','auditing','authorizing','awakening','backing','bargaining','bartering','becoming','building','calculating','campaigning','capturing','caring','carving','celebrating','certifying','changing','charging','cherishing','choosing','climbing','coding','commanding','communicating','compensating','competing','completing','concluding','confirming','connecting','considering','constructing','contemplating','controlling','cooking','coordinating','creating','curing','cutting','cycling','dancing','debating','deciding','defending','defining','delivering','designing','detecting','developing','directing','discussing','discovering','distributing','diving','dock','drawing','dreaming','driving','editing','educating','electing','emancipating','employing','empowering','enabling','endorsing','energizing','enjoying','equipping','etching','examining','existing','explaining','exploring','expressing','extracting','feeling','filming','financing','fishing','flying','forging','forgetting','functioning','funding','governing','greeting','growing','guarding','guiding','hacking','harvesting','healing','hearing','helping','highlighting','hiring','holding','honoring','hoping','hunting','identifying','illuminating','imagining','impacting','informing','innovating','inspecting','interpreting','investigating','jumping','keeping','labeling','landing','leading','leasing','learning','listening','living','loaning','lobbying','loving','maintaining','managing','manufacturing','marketing','measuring','mediating','medicating','mining','monitoring','molding','naming','navigating','negotiating','observing','offering','operating','orbiting','organizing','painting','performing','permitting','picking','planning','planting','playing','practicing','preparing','printing','processing','producing','programming','promoting','protecting','providing','publishing','racing','raising','reading','recalling','recording','recovering','refining','remembering','renting','rescuing','researching','resolving','respecting','rewarding','revitalizing','ruling','running','sailing','saving','sculpting','securing','seeing','selecting','selling','sensing','serving','shaping','shielding','shipping','singing','skating','skiing','smelling','solving','sponsoring','steering','stirring','studying','supporting','surfing','swimming','teaching','tracking','trading','trapping','traveling','treating','underlining','upholding','validating','viewing','voting','watching','winning','wishing','working','writing') | sort-object -unique))

# Create dropdown for Style
$styleLabel = New-Object System.Windows.Forms.Label
$styleLabel.Text = 'Style'
$styleLabel.AutoSize = $true
$styleLabel.Location = New-Object System.Drawing.Point(5, 110)
$styleCheckbox = New-Object System.Windows.Forms.CheckBox
$styleCheckbox.Location = New-Object System.Drawing.Point(180, 130)
$styleCheckbox.Text = 'Freeze',$styleLabel.Text
$styleCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($styleCheckbox)
$styleDropdown = New-Object System.Windows.Forms.ComboBox
$styleDropdown.Location = New-Object System.Drawing.Point(5, 130)
$styleDropdown.DropDownStyle = 'DropDown'
$styleDropdown.Size = New-Object System.Drawing.Size(165, 20)
$styledropdown.items.addrange((@('abstract art', 'abstract expressionism', 'academic art', 'acrylic paint', 'algorithmic art', 'antique', 'arte povera', 'art brut', 'art deco', 'art nouveau', 'augmented reality art', 'avant-garde', 'baroque', 'bauhaus', 'bio art', 'biological materials', 'bronze', 'canvas', 'caravaggism', 'ceramics', 'charcoal', 'classicism', 'clay', 'cobr', 'collage', 'color field painting', 'colored pencil', 'conceptual art', 'constructivism', 'contemporary art', 'cubism', 'cyber', 'cyber art', 'dadaism', 'de stijl', 'digital art', 'encaustic', 'environmental art', 'environmental materials', 'ephemeral materials', 'expressionism', 'fauvism', 'fiber', 'figurative art', 'fine art', 'found objects', 'futurism', 'futuristic', 'generative art', 'glass', 'gothic', 'graphite', 'graffiti', 'gouache paint', 'harlem renaissance', 'impressionism', 'ink', 'installation art', 'land art', 'leather', 'light', 'light art', 'lowbrow', 'lyrical abstraction', 'marble', 'medieval', 'metal', 'minimalism', 'modern', 'modern art', 'modernism', 'mosaic', 'naive art', 'nautical', 'neo-dada', 'neo-expressionism', 'neo-geo', 'neo-impressionism', 'neoclassicism', 'neon art', 'oil paint', 'op art', 'paper', 'pastel', 'performance art', 'performance elements', 'photorealism', 'picassian', 'plastics', 'pop art', 'porcelain', 'post-impressionism', 'post-internet', 'post-structuralism', 'postmodernism', 'precisionism', 'realism', 'renaissance', 'rembrandt lighting', 'relational aesthetics', 'retro', 'rococo', 'romantic', 'romanticism', 'sand', 'silverpoint', 'sound', 'sound art', 'steampunk', 'stone', 'street art', 'stuckism', 'superflat', 'suprematism', 'surrealism', 'symbolism', 'tachisme', 'tempera paint', 'textiles', 'toyism', 'transavantgarde', 'transavanguardia', 'video', 'video art', 'vintage', 'virtual reality art', 'watercolor paint', 'wood', 'young british artists', 'zero group') | foreach-object {'in the ' + $_ + ' style'} | sort-object -unique))

# Create dropdown for Media
$mediaLabel = New-Object System.Windows.Forms.Label
$mediaLabel.Text = 'Media'
$mediaLabel.AutoSize = $true
$mediaLabel.Location = New-Object System.Drawing.Point(5, 160)
$mediaCheckbox = New-Object System.Windows.Forms.CheckBox
$mediaCheckbox.Location = New-Object System.Drawing.Point(180, 180)
$mediaCheckbox.Text = 'Freeze',$mediaLabel.Text
$mediaCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($mediaCheckbox)
$mediaDropdown = New-Object System.Windows.Forms.ComboBox
$mediaDropdown.Location = New-Object System.Drawing.Point(5, 180)
$mediaDropdown.DropDownStyle = 'DropDown'
$mediaDropdown.Size = New-Object System.Drawing.Size(165, 20)
$mediadropdown.items.addrange((@('acrylic paint','alabaster','bamboo','beeswax','biologicals','bone','brass','bronze','car parts','cement','ceramic','chalk','charcoal','clay','colored pencils','copper','coral','crayon','dammar','diamonds','digital art','dirt','dye','egg tempera','encaustic','environmental materials','ephemeral materials','feathers','felt','fiber','film','found objects','fresco','garbage','gesso','glass','gold','gold leaf','gouache paint','granite','graphite','graphite powder','gypsum','ink','ivory','jade','lacquer','latex','leather','light','limestone','linoleum','lithography','magnet','mahogany','marble','metal','mosaic','mylar','oak','obsidian','oil paint','papyrus','paper','pastel','pastel pencils','pewter','plaster','plastics','plexiglass','polymer clay','porcelain','porphyry','resin','rubber','sand','silk','silver','silverpoint','slate','soapstone','sound','steel','stone','tempera paint','textiles','terracotta','utensils','vellum','velvet','venetian plaster','video','wax','wood') | foreach-object {'using ' + $_ + ' as artistic medium'} | sort-object -unique))

# Create dropdown for Color
$colorLabel = New-Object System.Windows.Forms.Label
$colorLabel.Text = 'Color'
$colorLabel.AutoSize = $true
$colorLabel.Location = New-Object System.Drawing.Point(5, 210)
$colorCheckbox = New-Object System.Windows.Forms.CheckBox
$colorCheckbox.Location = New-Object System.Drawing.Point(180, 230)
$colorCheckbox.Text = 'Freeze',$colorLabel.Text
$colorCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($colorCheckbox)
$colorDropdown = New-Object System.Windows.Forms.ComboBox
$colorDropdown.Location = New-Object System.Drawing.Point(5, 230)
$colorDropdown.DropDownStyle = 'DropDown'
$colorDropdown.Size = New-Object System.Drawing.Size(165, 20)
$colordropdown.items.addrange((@('analogous','arctic','autumn','berry','black and white','blue','bright','candy','chocolate','citrus','coffee','complementary','cool','dark','desaturated','desert','dull','earth tones','fire','floral','forest','galactic','green','ice','jewel tones','light','metallic','midnight','monochrome','muted','nature','neon','neutral','oceanic','pastel','primary','psychedelic','rainbow','red','rose','rose gold','rustic','saturated','secondary','sepia','smoky','split-complementary','spring','stormy','summer','sunrise','sunset','tertiary','tetradic','triadic','tropical','twilight','urban','vibrant','warm','wine','winter','yellow') | foreach-object {$_ + ' color palette'} | sort-object -unique))

# Create dropdown for Extra A
$extraALabel = New-Object System.Windows.Forms.Label
$extraALabel.Text = 'ExtraA'
$extraALabel.AutoSize = $true
$extraALabel.Location = New-Object System.Drawing.Point(5, 260)
$extraACheckbox = New-Object System.Windows.Forms.CheckBox
$extraACheckbox.Location = New-Object System.Drawing.Point(180, 280)
$extraACheckbox.Text = 'Freeze',$extraALabel.Text
$extraACheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($extraACheckbox)
$extraADropdown = New-Object System.Windows.Forms.ComboBox
$extraADropdown.Location = New-Object System.Drawing.Point(5, 280)
$extraADropdown.DropDownStyle = 'DropDown'
$extraADropdown.Size = New-Object System.Drawing.Size(165, 20)
$extraAdropdown.items.addrange((@('8k','bokeh','crisp','double exposure','dramatic','dreamy','ethereal','extreme close-up','extremely detailed','glossy finish','golden hour','grainy','hazy','high contrast','high dynamic range','high key','infrared','kodachrome','landscape','leading lines','lens flare','long exposure','low key','macro','matte finish','minimalist','moody','motion blur','panoramic view','portrait','reflections','rule of thirds','shallow depth of field','silhouette','smoke filled','smooth','soft focus','star trails','studio lighting','symmetry','textured','tilt-shifted','time-lapse','upside down','vignette','vintage','whimsical') | sort-object -unique))

# Create dropdown for Extra B
$extraBLabel = New-Object System.Windows.Forms.Label
$extraBLabel.Text = 'ExtraB'
$extraBLabel.AutoSize = $true
$extraBLabel.Location = New-Object System.Drawing.Point(5, 310)
$extraBCheckbox = New-Object System.Windows.Forms.CheckBox
$extraBCheckbox.Location = New-Object System.Drawing.Point(180, 330)
$extraBCheckbox.Text = 'Freeze',$extraBLabel.Text
$extraBCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($extraBCheckbox)
$extraBDropdown = New-Object System.Windows.Forms.ComboBox
$extraBDropdown.Location = New-Object System.Drawing.Point(5, 330)
$extraBDropdown.DropDownStyle = 'DropDown'
$extraBDropdown.Size = New-Object System.Drawing.Size(165, 20)
$extraBDropdown.items.addrange((@('3d rendered','aerial view','ambient light','backlit','birds eye view','bokehlicious','chiaroscuro','cross processing','cyberpunk','deep depth of field','dystopian','fantasy','film grain','fish-eye lens','flat lay','hard shadows','isometric','lens flare','light painting','macro shot','magical realism','negative space','panoramic','perspective','post-apocalyptic','rim light','sci-fi','silhouette','soft shadows','split toning','steampunk','sunburst','telephoto','utopian','wide angle','worms eye view','zoomed in','abstract','allegorical','aperture','blue color temperature','blue hour','brightness','candlelight','conceptual','contrast','dawn','depressive tone','diurnal','dusk','figurative','futuristic','geometric','glow','golden hour','high aperture','high iso','high shutter speed','historical','illumination','long focal length','low exposure','luminance','medium saturation','medium white balance','moonlight','mythological','narrative','neon','nightfall','nocturnal','opposite hue','organic','out of focus','radiance','retro','seasonal','starlight','symbolic','timeless','twilight') | sort-object -unique))

# Create dropdown for Extra C
$extraCLabel = New-Object System.Windows.Forms.Label
$extraCLabel.Text = 'ExtraC'
$extraCLabel.AutoSize = $true
$extraCLabel.Location = New-Object System.Drawing.Point(5, 360)
$extraCCheckbox = New-Object System.Windows.Forms.CheckBox
$extraCCheckbox.Location = New-Object System.Drawing.Point(180, 380)
$extraCCheckbox.Text = 'Freeze',$extraCLabel.Text
$extraCCheckbox.Size = New-Object System.Drawing.Size(65, 30)
$form.Controls.Add($extraCCheckbox)
$extraCDropdown = New-Object System.Windows.Forms.ComboBox
$extraCDropdown.Location = New-Object System.Drawing.Point(5, 380)
$extraCDropdown.DropDownStyle = 'DropDown'
$extraCDropdown.Size = New-Object System.Drawing.Size(165, 20)
$extracdropdown.items.addrange((@('aperture', 'bokeh', 'bracketing', 'camera body', 'compact camera', 'depth of field', 'digital camera', 'dslr', 'exposure', 'exposure compensation', 'film camera', 'filter', 'fish-eye lens', 'flash', 'focal length', 'hdr', 'histogram', 'image stabilization', 'iso', 'jpeg format', 'lens mount', 'light meter', 'long exposure', 'macro lens', 'medium format camera', 'mirrorless camera', 'monopod', 'neutral density filter', 'panorama', 'pixel', 'point-and-shoot camera', 'polarizing filter', 'prime lens', 'raw format', 'reflector', 'resolution', 'sensor', 'shutter speed', 'softbox', 'speedlight', 'telephoto lens', 'time-lapse', 'tripod', 'white balance', 'wide-angle lens', 'zoom lens') | sort-object -unique))

# Add event handlers to dropdowns
$subjectDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$actionDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$styleDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$mediaDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$colorDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$extraADropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$extraBDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$extraCDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)

$dropdowns = @($subjectDropdown, $actionDropdown, $styleDropdown, $mediaDropdown, $colorDropdown, $extraADropdown, $extraBDropdown, $extraCDropdown)

foreach ($dropdown in $dropdowns) {
    $dropdown.Add_SelectedIndexChanged({
        $selectedWords = @()
        foreach ($dd in $dropdowns) {
            if ($dd.SelectedItem) { $selectedWords += $dd.SelectedItem }
        }
        $clipboardText = [string]::Join(', ', $selectedWords)
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
    $confirmation = [System.Windows.Forms.MessageBox]::Show('Are you sure you want to generate a random prompt? This will overwrite any chosen dropdowns!','Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
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
        if(!$extraACheckbox.Checked){
            $randomWord = $extraADropdown.Items | Get-Random
            $selectedWords += $randomWord
            $extraADropdown.SelectedItem = $randomWord
        }
        if(!$extraBCheckbox.Checked){
            $randomWord = $extraBDropdown.Items | Get-Random
            $selectedWords += $randomWord
            $extraBDropdown.SelectedItem = $randomWord
        }
        if(!$extraCCheckbox.Checked){
            $randomWord = $extraCDropdown.Items | Get-Random
            $selectedWords += $randomWord
            $extraCDropdown.SelectedItem = $randomWord
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
    if ($copiedTextbox.Text -ne $null -and $copiedTextbox.Text -ne '') {
        Get-Clipboard
    }
    else {
        [System.Windows.Forms.MessageBox]::Show('Clipboard is empty. Please select at least one item.','Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $baseURL = 'https://designer.microsoft.com/image-creator?p='

    # Open the website with the clipboard text
    $formattedText = Format-ForURL -text $copiedTextbox.Text
    if ($formattedText) {
        $url = $baseURL + $formattedText + '&form=NTPCH1&refig=dd8bc5ce95bf495a89a1b6c447914a00&pc=U531&adppc=EDGEDBB&sp=2&lq=0&qs=PN&sk=PN1&sc=8-0&cvid=dd8bc5ce95bf495a89a1b6c447914a00&showconv=1&sendquery=1'
        Start-Process $url
    } 
    else {
        [System.Windows.Forms.MessageBox]::Show('Clipboard is empty. Please select at least one item.','Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }

    # Log the clipboard text to a file
    if ($formattedText) {
        $logFile = 'prompt.log'
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = '$timestamp - $clipboardText'
        Add-Content -Path $logFile -Value $logEntry
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
        [System.Windows.Forms.MessageBox]::Show('Clipboard is empty. Please select at least one item.','Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $baseURL = 'https://www.bing.com/search?q='

    # Open the website with the clipboard text
    $formattedText = Format-ForURL -text $copiedTextbox.Text
    if ($formattedText) {
        $url = $baseURL + 'an image of ' + $formattedText + '&form=NTPCH1&refig=dd8bc5ce95bf495a89a1b6c447914a00&pc=U531&adppc=EDGEDBB&sp=2&lq=0&qs=PN&sk=PN1&sc=8-0&cvid=dd8bc5ce95bf495a89a1b6c447914a00&showconv=1&sendquery=1&dissrchswrite=1'
        Start-Process $url
    } else {
        [System.Windows.Forms.MessageBox]::Show('Clipboard is empty. Please select at least one item.','Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    
    # Log the clipboard text to a file
    if ($formattedText) {
        $logFile = 'prompt.log'
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "$timestamp - $clipboardText"
        Add-Content -Path $logFile -Value $logEntry
    }
})

# Create a button to reset the dropdown boxes
$resetButton = New-Object System.Windows.Forms.Button
$resetButton.Location = New-Object System.Drawing.Point(400, 30)
$resetButton.Size = New-Object System.Drawing.Size(175, 370)
$resetButton.Text = 'Reset Dropdowns and Clear Clipboard'

# Add click event to the reset dropdown boxes button
$resetButton.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show('Are you sure you want to reset the dropdown boxes and clear the clipboard?','Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmation -eq 'Yes') {
        if ($subjectDropdown.Items.Count -gt 0) {$subjectDropdown.SelectedIndex = -1}
        if ($actionDropdown.Items.Count -gt 0) {$actionDropdown.SelectedIndex = -1}
        if ($styleDropdown.Items.Count -gt 0) {$styleDropdown.SelectedIndex = -1}
        if ($mediaDropdown.Items.Count -gt 0) {$mediaDropdown.SelectedIndex = -1}
        if ($colorDropdown.Items.Count -gt 0) {$colorDropdown.SelectedIndex = -1}
        if ($extraADropdown.Items.Count -gt 0) {$extraADropdown.SelectedIndex = -1}
        if ($extraBDropdown.Items.Count -gt 0) {$extraBDropdown.SelectedIndex = -1}
        if ($extraCDropdown.Items.Count -gt 0) {$extraCDropdown.SelectedIndex = -1}

        # Reset the checkboxes
        $subjectCheckbox.Checked = $false
        $actionCheckbox.Checked = $false
        $styleCheckbox.Checked = $false
        $mediaCheckbox.Checked = $false
        $colorCheckbox.Checked = $false
        $extraACheckbox.Checked = $false
        $extraBCheckbox.Checked = $false
        $extraCCheckbox.Checked = $false

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
$form.Controls.AddRange(@($subjectLabel, $subjectDropdown, $subjectLabelText, $actionLabel, $actionDropdown, $styleLabel, $styleDropdown, $mediaLabel, $mediaDropdown, $colorLabel, $colorDropdown, $extraALabel, $extraADropdown, $extraBLabel, $extraBDropdown, $extraCLabel, $extraCDropdown, $copybutton, $copiedLabel, $copiedTextbox))
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