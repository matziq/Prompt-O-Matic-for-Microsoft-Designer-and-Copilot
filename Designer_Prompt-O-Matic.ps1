<#
Author: John Rea
Date: April 25, 2024
Version: 1.3
#>

#Adds the System.Windows.Forms assembly.
Add-Type -AssemblyName System.Windows.Forms

# Clear the clipboard on launch
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Clipboard]::Clear()

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Prompt-O-Matic for Microsoft Designer'
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.AutoSize = $true

# Create label and textbox to display copied text
$copiedLabel = New-Object System.Windows.Forms.Label
$copiedLabel.Text = 'Current Prompt in the Clipboard'
$copiedLabel.AutoSize = $true
$copiedLabel.Location = New-Object System.Drawing.Point(10, 410)
$copiedTextbox = New-Object System.Windows.Forms.TextBox
$copiedTextbox.Location = New-Object System.Drawing.Point(10, 430)
$copiedTextbox.ReadOnly = $true
$copiedTextbox.Size = New-Object System.Drawing.Size(550, 50)
$copiedTextbox.Multiline = $true
$copiedTextbox.ScrollBars = 'Vertical'

# Create filter textboxes for each dropdown
$subjectFilterTextbox = New-Object System.Windows.Forms.TextBox
$subjectFilterTextbox.Location = New-Object System.Drawing.Point(200, 10)
$actionFilterTextbox = New-Object System.Windows.Forms.TextBox
$actionFilterTextbox.Location = New-Object System.Drawing.Point(200, 30)
$styleFilterTextbox = New-Object System.Windows.Forms.TextBox
$styleFilterTextbox.Location = New-Object System.Drawing.Point(200, 50)
$mediaFilterTextbox = New-Object System.Windows.Forms.TextBox
$mediaFilterTextbox.Location = New-Object System.Drawing.Point(200, 70)
$colorFilterTextbox = New-Object System.Windows.Forms.TextBox
$colorFilterTextbox.Location = New-Object System.Drawing.Point(200, 90)
$extraAFilterTextbox = New-Object System.Windows.Forms.TextBox
$extraAFilterTextbox.Location = New-Object System.Drawing.Point(200, 110)
$extraBFilterTextbox = New-Object System.Windows.Forms.TextBox
$extraBFilterTextbox.Location = New-Object System.Drawing.Point(200, 130)
$extraCFilterTextbox = New-Object System.Windows.Forms.TextBox
$extraCFilterTextbox.Location = New-Object System.Drawing.Point(200, 150)

# Create dropdown for Subject
$subjectLabel = New-Object System.Windows.Forms.Label
$subjectLabel.Text = 'Subject of Artwork:'
$subjectLabel.AutoSize = $true
$subjectLabel.Location = New-Object System.Drawing.Point(10, 10)
$subjectDropdown = New-Object System.Windows.Forms.ComboBox
$subjectDropdown.Location = New-Object System.Drawing.Point(10, 30)
$subjectdropdown.items.addrange((@('airplane', 'alligator', 'archipelago', 'backyard', 'balloon', 'bay', 'beach', 'bear', 'bedroom', 'bicycle', 'book', 'bridge', 'butterfly', 'camera', 'cape', 'castle', 'cat', 'cave', 'city', 'cliff', 'cloud', 'computer', 'coral', 'crocodile', 'den', 'desert', 'dolphin', 'eagle', 'elephant', 'flower', 'forest', 'fountain', 'frog', 'galaxy', 'garden', 'garage', 'giraffe', 'glacier', 'guitar', 'harbor', 'helicopter', 'hill', 'hippopotamus', 'horse', 'iceberg', 'island', 'jellyfish', 'kangaroo', 'kitchen', 'kite', 'koala', 'lamp', 'lighthouse', 'lion', 'living room', 'meadow', 'monkey', 'moon', 'mountain', 'museum', 'ocean', 'octopus', 'owl', 'panda', 'parrot', 'patio', 'peacock', 'pebble', 'penguin', 'peninsula', 'piano', 'pond', 'pyramid', 'rabbit', 'river', 'rose', 'sand', 'seashell', 'shark', 'skyscraper', 'snake', 'squirrel', 'star', 'statue', 'study', 'sunset', 'telescope', 'temple', 'tesla', 'tiger', 'train', 'tree', 'truck', 'turtle', 'valley', 'volcano', 'waterfall', 'whale', 'wolf', 'zebra') | ForEach-Object {'a ' + $_} | sort-object -Unique))

# Create dropdown for Action
$actionLabel = New-Object System.Windows.Forms.Label
$actionLabel.Text = 'Action Verb:'
$actionLabel.AutoSize = $true
$actionLabel.Location = New-Object System.Drawing.Point(10, 60)
$actionDropdown = New-Object System.Windows.Forms.ComboBox
$actionDropdown.Location = New-Object System.Drawing.Point(10, 80)
$actionDropdown.Items.AddRange((@('acting', 'admiring', 'advertising', 'aiding', 'allowing', 'analyzing', 'appreciating', 'appointing', 'arguing', 'arranging', 'aspiring', 'assembling', 'assisting', 'auditing', 'authorizing', 'backing', 'bargaining', 'bartering', 'becoming', 'building', 'calculating', 'campaigning', 'capturing', 'caring', 'carving', 'celebrating', 'certifying', 'changing', 'charging', 'cherishing', 'choosing', 'climbing', 'coding', 'commanding', 'communicating', 'compensating', 'competing', 'completing', 'concluding', 'confirming', 'connecting', 'considering', 'constructing', 'contemplating', 'controlling', 'cooking', 'coordinating', 'creating', 'curing', 'cutting', 'cycling', 'dancing', 'debating', 'deciding', 'defending', 'defining', 'delivering', 'designing', 'detecting', 'developing', 'directing', 'discussing', 'discovering', 'distributing', 'diving', 'dock', 'drawing', 'dreaming', 'driving', 'editing', 'educating', 'electing', 'emancipating', 'employing', 'empowering', 'enabling', 'endorsing', 'enjoying', 'equipping', 'etching', 'examining', 'existing', 'explaining', 'exploring', 'expressing', 'extracting', 'feeling', 'filming', 'financing', 'fishing', 'flying', 'forging', 'forgetting', 'functioning', 'funding', 'governing', 'greeting', 'growing', 'guarding', 'guiding', 'hacking', 'harvesting', 'healing', 'hearing', 'helping', 'highlighting', 'hiring', 'holding', 'honoring', 'hoping', 'hunting', 'identifying', 'imagining', 'informing', 'innovating', 'inspecting', 'interpreting', 'investigating', 'jumping', 'keeping', 'labeling', 'landing', 'leading', 'leasing', 'learning', 'listening', 'living', 'loaning', 'lobbying', 'loving', 'maintaining', 'managing', 'manufacturing', 'marketing', 'measuring', 'mediating', 'medicating', 'mining', 'monitoring', 'molding', 'naming', 'navigating', 'negotiating', 'observing', 'offering', 'operating', 'orbiting', 'organizing', 'painting', 'performing', 'permitting', 'picking', 'planning', 'planting', 'playing', 'practicing', 'preparing', 'printing', 'processing', 'producing', 'programming', 'promoting', 'protecting', 'providing', 'publishing', 'racing', 'raising', 'reading', 'recalling', 'recording', 'recovering', 'refining', 'remembering', 'renting', 'rescuing', 'researching', 'resolving', 'respecting', 'rewarding', 'ruling', 'running', 'sailing', 'saving', 'sculpting', 'securing', 'seeing', 'selecting', 'selling', 'sensing', 'serving', 'shaping', 'shielding', 'shipping', 'singing', 'skating', 'skiing', 'smelling', 'solving', 'sponsoring', 'steering', 'studying', 'supporting', 'surfing', 'swimming', 'teaching', 'tracking', 'trading', 'trapping', 'traveling', 'treating', 'underlining', 'upholding', 'validating', 'viewing', 'voting', 'watching', 'winning', 'wishing', 'working', 'writing') | Sort-Object -Unique))

# Create dropdown for Style
$styleLabel = New-Object System.Windows.Forms.Label
$styleLabel.Text = 'Artistic Style:'
$styleLabel.AutoSize = $true
$styleLabel.Location = New-Object System.Drawing.Point(10, 110)
$styleDropdown = New-Object System.Windows.Forms.ComboBox
$styleDropdown.Location = New-Object System.Drawing.Point(10, 130)
$styleDropdown.Items.AddRange((@('abstract art', 'academic art', 'acrylic paint', 'antique', 'arte povera', 'art brut', 'art deco', 'art nouveau', 'avant-garde', 'bauhaus', 'biological materials', 'bronze', 'canvas', 'caravaggism', 'ceramics', 'charcoal', 'classicism', 'clay', 'cobr', 'collage', 'color field painting', 'colored pencil', 'conceptual art', 'constructivism', 'contemporary art', 'cubism', 'cyber', 'dadaism', 'de stijl', 'digital art', 'encaustic', 'environmental materials', 'ephemeral materials', 'expressionism', 'fauvism', 'fiber', 'figurative art', 'fine art', 'found objects', 'futurism', 'futuristic', 'glass', 'gothic', 'graphite', 'gouache paint', 'harlem renaissance', 'ink', 'installation art', 'land art', 'leather', 'light', 'lyrical abstraction', 'marble', 'medieval', 'metal', 'minimalism', 'modern', 'modern art', 'mosaic', 'naive art', 'nautical', 'neo-dada', 'neo-expressionism', 'neo-impressionism', 'neoclassicism', 'neon art', 'oil paint', 'op art', 'paper', 'pastel', 'performance elements', 'photorealism', 'picassian', 'plastics', 'pop art', 'porcelain', 'post-impressionism', 'precisionism', 'realism', 'renaissance', 'rembrandt lighting', 'retro', 'rococo', 'romantic', 'sand', 'silverpoint', 'sound', 'steampunk', 'stone', 'street art', 'suprematism', 'surrealism', 'symbolism', 'tachisme', 'tempera paint', 'textiles', 'transavanguardia', 'video', 'vintage', 'watercolor paint', 'wood', 'zero group') | ForEach-Object {'in the ' + $_ + ' style'} | Sort-Object -Unique))

# Create dropdown for Media
$mediaLabel = New-Object System.Windows.Forms.Label
$mediaLabel.Text = 'Type of Media:'
$mediaLabel.AutoSize = $true
$mediaLabel.Location = New-Object System.Drawing.Point(10, 160)
$mediaDropdown = New-Object System.Windows.Forms.ComboBox
$mediaDropdown.Location = New-Object System.Drawing.Point(10, 180)
$mediaDropdown.Items.AddRange((@('acrylic paint', 'alabaster', 'bamboo', 'beeswax', 'biologicals', 'bone', 'brass', 'bronze', 'car parts', 'cement', 'ceramic', 'chalk', 'charcoal', 'clay', 'colored pencils', 'copper', 'coral', 'crayon', 'dammar', 'diamonds', 'digital art', 'dirt', 'dye', 'egg tempera', 'encaustic', 'environmental materials', 'ephemeral materials', 'feathers', 'felt', 'fiber', 'found objects', 'fresco', 'garbage', 'gesso', 'glass', 'gold', 'gold leaf', 'gouache paint', 'granite', 'graphite', 'graphite powder', 'gypsum', 'ink', 'ivory', 'jade', 'lacquer', 'latex', 'leather', 'light', 'limestone', 'linoleum', 'lithography', 'magnet', 'mahogany', 'marble', 'metal', 'mosaic', 'mylar', 'oak', 'obsidian', 'oil paint', 'papyrus', 'paper', 'pastel', 'pastel pencils', 'pewter', 'plaster', 'plastics', 'plexiglass', 'polymer clay', 'porcelain', 'porphyry', 'resin', 'rubber', 'sand', 'silk', 'silver', 'silverpoint', 'slate', 'soapstone', 'sound', 'steel', 'stone', 'tempera paint', 'textiles', 'terracotta', 'utensils', 'vellum', 'velvet', 'venetian plaster', 'video', 'wax', 'wood') | ForEach-Object {$_ + ' as artistic medium'} | Sort-Object -Unique))

# Create dropdown for Color
$colorLabel = New-Object System.Windows.Forms.Label
$colorLabel.Text = 'Type of Color Palette:'
$colorLabel.AutoSize = $true
$colorLabel.Location = New-Object System.Drawing.Point(10, 210)
$colorDropdown = New-Object System.Windows.Forms.ComboBox
$colorDropdown.Location = New-Object System.Drawing.Point(10, 230)
$colorDropdown.Items.AddRange((@('analogous', 'arctic', 'autumn', 'berry', 'black and white', 'blue', 'bright', 'candy', 'chocolate', 'citrus', 'coffee', 'complementary', 'cool', 'dark', 'desaturated', 'desert', 'dull', 'earth tones', 'fire', 'floral', 'forest', 'galactic', 'green', 'ice', 'jewel tones', 'light', 'metallic', 'midnight', 'monochrome', 'muted', 'nature', 'neon', 'neutral', 'oceanic', 'pastel', 'primary', 'psychedelic', 'rainbow', 'red', 'rose', 'rose gold', 'rustic', 'saturated', 'secondary', 'sepia', 'smoky', 'split-complementary', 'spring', 'stormy', 'summer', 'sunrise', 'sunset', 'tertiary', 'tetradic', 'triadic', 'tropical', 'twilight', 'urban', 'vibrant', 'warm', 'wine', 'winter', 'yellow') | ForEach-Object {'with a ' + $_ + ' color palette'} | Sort-Object -Unique))

# Create dropdown for Extra A
$extraALabel = New-Object System.Windows.Forms.Label
$extraALabel.Text = 'Extra A:'
$extraALabel.AutoSize = $true
$extraALabel.Location = New-Object System.Drawing.Point(10, 260)
$extraADropdown = New-Object System.Windows.Forms.ComboBox
$extraADropdown.Location = New-Object System.Drawing.Point(10, 280)
$extraADropdown.Items.AddRange((@('8k', 'bokeh', 'crisp', 'double exposure', 'dramatic', 'dreamy', 'ethereal', 'extreme close-up', 'extremely detailed', 'glossy finish', 'golden hour', 'grainy', 'hazy', 'high contrast', 'high dynamic range', 'high key', 'infrared', 'kodachrome', 'landscape', 'leading lines', 'lens flare', 'long exposure', 'low key', 'macro', 'matte finish', 'minimalist', 'moody', 'motion blur', 'panoramic view', 'portrait', 'reflections', 'rule of thirds', 'shallow depth of field', 'silhouette', 'smoke filled', 'smooth', 'soft focus', 'star trails', 'studio lighting', 'symmetry', 'textured', 'tilt-shifted', 'time-lapse', 'upside down', 'vignette', 'vintage', 'whimsical') | Sort-Object -Unique))

# Create dropdown for Extra B
$extraBLabel = New-Object System.Windows.Forms.Label
$extraBLabel.Text = 'Extra B:'
$extraBLabel.AutoSize = $true
$extraBLabel.Location = New-Object System.Drawing.Point(10, 310)
$extraBDropdown = New-Object System.Windows.Forms.ComboBox
$extraBDropdown.Location = New-Object System.Drawing.Point(10, 330)
$extraBDropdown.Items.AddRange((@('3d rendered', 'aerial view', 'ambient light', 'backlit', 'birds eye view', 'bokehlicious', 'chiaroscuro', 'cross processing', 'cyberpunk', 'deep depth of field', 'dystopian', 'fantasy', 'film grain', 'fish-eye lens', 'flat lay', 'hard shadows', 'isometric', 'lens flare', 'light painting', 'macro shot', 'magical realism', 'negative space', 'panoramic', 'perspective', 'post-apocalyptic', 'rim light', 'sci-fi', 'silhouette', 'soft shadows', 'split toning', 'steampunk', 'sunburst', 'telephoto', 'utopian', 'wide angle', 'worms eye view', 'zoomed in') | Sort-Object -Unique))

# Create dropdown for Extra C
$extraCLabel = New-Object System.Windows.Forms.Label
$extraCLabel.Text = 'Extra C:'
$extraCLabel.AutoSize = $true
$extraCLabel.Location = New-Object System.Drawing.Point(10, 360)
$extraCDropdown = New-Object System.Windows.Forms.ComboBox
$extraCDropdown.Location = New-Object System.Drawing.Point(10, 380)
$extraCDropdown.Items.AddRange((@('abstract', 'allegorical', 'aperture', 'blue color temperature', 'blue hour', 'brightness', 'candlelight', 'conceptual', 'contrast', 'dawn', 'depressive tone', 'diurnal', 'dusk', 'figurative', 'futuristic', 'geometric', 'glow', 'golden hour', 'high aperture', 'high iso', 'high shutter speed', 'historical', 'illumination', 'long focal length', 'low exposure', 'luminance', 'medium saturation', 'medium white balance', 'moonlight', 'mythological', 'narrative', 'neon', 'nightfall', 'nocturnal', 'opposite hue', 'organic', 'out of focus', 'radiance', 'retro', 'seasonal', 'starlight', 'symbolic', 'timeless', 'twilight') | Sort-Object -Unique))

# Enable auto-complete mode for dropdowns
$subjectDropdown.AutoCompleteMode = 'SuggestAppend'
$subjectDropdown.AutoCompleteSource = 'ListItems'
$subjectDropdown.DropDownStyle = 'DropDown'
$actionDropdown.AutoCompleteMode = 'SuggestAppend'
$actionDropdown.AutoCompleteSource = 'ListItems'
$actionDropdown.DropDownStyle = 'DropDown'
$styleDropdown.AutoCompleteMode = 'SuggestAppend'
$styleDropdown.AutoCompleteSource = 'ListItems'
$styleDropdown.DropDownStyle = 'DropDown'
$mediaDropdown.AutoCompleteMode = 'SuggestAppend'
$mediaDropdown.AutoCompleteSource = 'ListItems'
$mediaDropdown.DropDownStyle = 'DropDown'
$colorDropdown.AutoCompleteMode = 'SuggestAppend'
$colorDropdown.AutoCompleteSource = 'ListItems'
$colorDropdown.DropDownStyle = 'DropDown'
$extraADropdown.AutoCompleteMode = 'SuggestAppend'
$extraADropdown.AutoCompleteSource = 'ListItems'
$extraADropdown.DropDownStyle = 'DropDown'
$extraBDropdown.AutoCompleteMode = 'SuggestAppend'
$extraBDropdown.AutoCompleteSource = 'ListItems'
$extraBDropdown.DropDownStyle = 'DropDown'
$extraCDropdown.AutoCompleteMode = 'SuggestAppend'
$extraCDropdown.AutoCompleteSource = 'ListItems'
$extraCDropdown.DropDownStyle = 'DropDown'

# Create button to copy to clipboard
$button = New-Object System.Windows.Forms.Button
$button.Size = New-Object System.Drawing.Size(75, 370)
$button.Text = 'Copy Selections to the Clipboard to form a Prompt'
$button.Location = New-Object System.Drawing.Point(140, 30)
$button.Enabled = $false

# Enable button only when all dropdowns have a selection
$subjectDropdown.add_SelectedIndexChanged({
    $button.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem
})

$actionDropdown.add_SelectedIndexChanged({
    $button.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem
})

$styleDropdown.add_SelectedIndexChanged({
    $button.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem
})

$mediaDropdown.add_SelectedIndexChanged({
    $button.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem
})

$colorDropdown.add_SelectedIndexChanged({
    $button.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem
})

$extraADropdown.add_SelectedIndexChanged({
    $button.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem
})

$extraBDropdown.add_SelectedIndexChanged({
    $button.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem
})

$extraCDropdown.add_SelectedIndexChanged({
    $button.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem
})
$button.Add_Click({
    $selectedWords = @()
    if ($subjectDropdown.SelectedItem) { $selectedWords += $subjectDropdown.SelectedItem }
    if ($actionDropdown.SelectedItem) { $selectedWords += $actionDropdown.SelectedItem }
    if ($styleDropdown.SelectedItem) { $selectedWords += $styleDropdown.SelectedItem }
    if ($mediaDropdown.SelectedItem) { $selectedWords += $mediaDropdown.SelectedItem }
    if ($colorDropdown.SelectedItem) { $selectedWords += $colorDropdown.SelectedItem }
    if ($extraADropdown.SelectedItem) { $selectedWords += $extraADropdown.SelectedItem }
    if ($extraBDropdown.SelectedItem) { $selectedWords += $extraBDropdown.SelectedItem }
    if ($extraCDropdown.SelectedItem) { $selectedWords += $extraCDropdown.SelectedItem }
    if ($selectedWords.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one item before clicking the button.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $clipboardText = [string]::Join(', ', $selectedWords)
    $clipboardText = "Create an image of " + $clipboardText #.Substring(0,1).ToUpper()+$clipboardText.Substring(1) + "."
    Set-Clipboard -Value $clipboardText
    if ($clipboardText) {
        $copiedTextbox.Text = $clipboardText
    }
})

# Create "Random" button
$randomButton = New-Object System.Windows.Forms.Button
$randomButton.Text = 'Random Prompt Generator'
$randomButton.Location = New-Object System.Drawing.Point(230, 30)
$randomButton.Size = New-Object System.Drawing.Size(150, 180)
$randomButton.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to generate a random prompt?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmation -eq "Yes") {
        $selectedWords = @()
        $selectedWords += $subjectDropdown.Items | Get-Random
        $selectedWords += $actionDropdown.Items | Get-Random
        $selectedWords += $styleDropdown.Items | Get-Random
        $selectedWords += $mediaDropdown.Items | Get-Random
        $selectedWords += $colorDropdown.Items | Get-Random
        $selectedWords += $extraADropdown.Items | Get-Random
        $selectedWords += $extraBDropdown.Items | Get-Random
        $selectedWords += $extraCDropdown.Items | Get-Random
        $clipboardText = "Create an image of " + [string]::Join(', ', $selectedWords)
        if ($clipboardText) {
            $clipboardText = $clipboardText.Substring(0,1).ToUpper()+$clipboardText.Substring(1) + "."
            try {
                Set-Clipboard -Value $clipboardText
            } catch {
                [System.Windows.Forms.MessageBox]::Show("An error occurred while setting the clipboard: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            $copiedTextbox.Text = $clipboardText
        }

        # Update the dropdown fields in the GUI
        $subjectDropdown.SelectedItem = $selectedWords[0]
        $actionDropdown.SelectedItem = $selectedWords[1]
        $styleDropdown.SelectedItem = $selectedWords[2]
        $mediaDropdown.SelectedItem = $selectedWords[3]
        $colorDropdown.SelectedItem = $selectedWords[4]
        $extraADropdown.SelectedItem = $selectedWords[5]
        $extraBDropdown.SelectedItem = $selectedWords[6]
        $extraCDropdown.SelectedItem = $selectedWords[7]
    }
})

# Set the copied text to the joined selected words
if ($selectedWords -ne $null) {
    $copiedTextbox.Text = $selectedWords -join ' '
} else {
    if ($copiedTextbox -ne $null) {
        $copiedTextbox.Text = ""
    }
}

# Add a button to open the website with the generated text
$webButton = New-Object System.Windows.Forms.Button
$webButton.Text = 'Open Current Prompt in Designer (or paste into Copilot)'
$webButton.Location = New-Object System.Drawing.Point(230, 220)
$webButton.Size = New-Object System.Drawing.Size(150, 180)

# Disable the button initially
$webButton.Enabled = $copiedTextbox.Text -ne ""

$webButton.Add_Click({
    # Get the current clipboard text
    $clipboardText = Get-Clipboard

    if ($clipboardText) {
        $copiedTextbox.Text = $clipboardText
    } else {
        [System.Windows.Forms.MessageBox]::Show("Clipboard is empty. Please select at least one item.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Define the base URL for the website
    $baseURL = "https://designer.microsoft.com/image-creator?p="
    # Create a function to format the text for the URL
    function Format-ForURL ($text) {
        try {
            $text = $text.Replace(' ', '+')
            return $text
        } catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred while formatting the text: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    # Open the website with the clipboard text
    $formattedText = Format-ForURL -text $copiedTextbox.Text
    if ($formattedText) {
        $url = $baseURL + $formattedText
        Start-Process $url
    } else {
        [System.Windows.Forms.MessageBox]::Show("Clipboard is empty. Please select at least one item.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }

    # Log the clipboard text to a file
    if ($formattedText) {
        $logFile = "prompt.log"
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timestamp - $clipboardText"
        Add-Content -Path $logFile -Value $logEntry
    }
})

# Add an event handler to enable/disable the button based on the clipboard text
$copiedTextbox.add_TextChanged({
    $webButton.Enabled = $copiedTextbox.Text -ne ""
})

# Add the selected items from the Dropdowns to the selectedWords array
if ($subjectDropdown.Text) { $selectedWords += $subjectDropdown.Text }
if ($actionDropdown.Text) { $selectedWords += $actionDropdown.Text }
if ($styleDropdown.Text) { $selectedWords += $styleDropdown.Text }
if ($mediaDropdown.Text) { $selectedWords += $mediaDropdown.Text }
if ($colorDropdown.Text) { $selectedWords += $colorDropdown.Text }
if ($extraADropdown.Text) { $selectedWords += $extraADropdown.Text }
if ($extraBDropdown.Text) { $selectedWords += $extraBDropdown.Text }
if ($extraCDropdown.Text) { $selectedWords += $extraCDropdown.Text }

# Create a button to reset the dropdown boxes
$resetButton = New-Object System.Windows.Forms.Button
$resetButton.Location = New-Object System.Drawing.Point(400, 30)
$resetButton.Size = New-Object System.Drawing.Size(150, 180)
$resetButton.Text = "Reset Dropdowns"

# Add click event to the reset dropdown boxes button
$resetButton.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to reset the dropdown boxes?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmation -eq "Yes") {
        if ($subjectDropdown.Items.Count -gt 0) {
            $subjectDropdown.SelectedIndex = -1
        }
        if ($actionDropdown.Items.Count -gt 0) {
            $actionDropdown.SelectedIndex = -1
        }
        if ($styleDropdown.Items.Count -gt 0) {
            $styleDropdown.SelectedIndex = -1
        }
        if ($mediaDropdown.Items.Count -gt 0) {
            $mediaDropdown.SelectedIndex = -1
        }
        if ($colorDropdown.Items.Count -gt 0) {
            $colorDropdown.SelectedIndex = -1
        }
        if ($extraADropdown.Items.Count -gt 0) {
            $extraADropdown.SelectedIndex = -1
        }
        if ($extraBDropdown.Items.Count -gt 0) {
            $extraBDropdown.SelectedIndex = -1
        }
        if ($extraCDropdown.Items.Count -gt 0) {
            $extraCDropdown.SelectedIndex = -1
        }
    }
})

# Create a button to clear the clipboard
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Location = New-Object System.Drawing.Point(400, 220)
$clearButton.Size = New-Object System.Drawing.Size(150, 180)
$clearButton.Text = "Clear Clipboard"

# Disable the button initially
$clearButton.Enabled = [System.Windows.Forms.Clipboard]::ContainsText()

# Add click event to the clear button
$clearButton.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to clear the clipboard?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmation -eq "Yes") {
        try {
            if ([System.Windows.Forms.Clipboard]::ContainsText()) {
                [System.Windows.Forms.Clipboard]::Clear()
                $copiedTextbox.Text = ""
                $clearButton.Enabled = $false
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred while clearing the clipboard: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Add event handler to enable/disable the clear button based on clipboard content
$copiedTextbox.add_TextChanged({
    $clearButton.Enabled = [System.Windows.Forms.Clipboard]::ContainsText()
})

# Add event handler to enable/disable the reset button based on dropdown selection
$subjectDropdown.add_SelectedIndexChanged({
    $resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem
})

$actionDropdown.add_SelectedIndexChanged({
    $resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem
})

$styleDropdown.add_SelectedIndexChanged({
    $resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem
})

$mediaDropdown.add_SelectedIndexChanged({
    $resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem
})

$colorDropdown.add_SelectedIndexChanged({
    $resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem
})

$extraADropdown.add_SelectedIndexChanged({
    $resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem
})

$extraBDropdown.add_SelectedIndexChanged({
    $resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem
})

$extraCDropdown.add_SelectedIndexChanged({
    $resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem
})

# Add the reset dropdown boxes button to the form
$form.Controls.Add($resetButton)

# Add the web button to the form controls
$form.Controls.Add($webButton)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

# Add the random button to the form controls
$form.Controls.Add($randomButton)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

# Add the label to the form controls
$form.Controls.Add($copiedLabel)

# Add the clear button to the form controls
$form.Controls.Add($clearButton)

# Add controls to form
$form.Controls.AddRange(@($subjectLabel, $subjectDropdown, $actionLabel, $actionDropdown, $styleLabel, $styleDropdown, $mediaLabel, $mediaDropdown, $colorLabel, $colorDropdown, $extraALabel, $extraADropdown, $extraBLabel, $extraBDropdown,$extraCLabel, $extraCDropdown,$button, $copiedLabel, $copiedTextbox))
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

# Show form
$form.ShowDialog()
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

# SIG # Begin signature block
# MIIJCQYJKoZIhvcNAQcCoIII+jCCCPYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUznP/xRirwiC4RwehpJl6bv3A
# jWSgggZqMIIGZjCCBU6gAwIBAgITJAAAAUPgt24PVLctvwAAAAABQzANBgkqhkiG
# 9w0BAQsFADBRMRMwEQYKCZImiZPyLGQBGRYDb3JnMRwwGgYKCZImiZPyLGQBGRYM
# bWVtY29uZmlnbWdyMRwwGgYDVQQDExNtZW1jb25maWdtZ3ItQ00xLUNBMB4XDTI0
# MDQxNzE0MzI1OVoXDTI2MDQxNzE0MzI1OVowEzERMA8GA1UEAxMIbGFiYWRtaW4w
# ggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDv2ENL7267CRUs4KG//Mnf
# UQj9JKkVF0SZOLZFWRxkc7md06Ie3sPujBJ8WQDRZGex6DvmucHPDGKC8yxQRnbc
# /i1KH9k2UX0U3baag0CWSGoBgZyBuTMY5/yIFiQEJsEue6W51kN9LyDgvvbMNLi6
# n3Typ5e4VgfqlQay6HtJSeZet5s2BQMyOt3LjBHs/SSpTCMzk2/DrR5lf4gIfNb/
# xI6loNzlhTZ1WCurMES8UL0jRyvsqYsEJ+1fCD8tXXVjxWxOdUXAUjpPtkJrwhBC
# PUIyXIAF/PTvYUVu1KwYu7MKuH0Jav0lTH9pFWndDRC7/3bsXHREf53bgtmFXgep
# AgMBAAGjggNzMIIDbzA8BgkrBgEEAYI3FQcELzAtBiUrBgEEAYI3FQjPpDWH0vEw
# hPWJNoOEnwCGjaALXIHCphqGrf1fAgFsAgEAMBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MA4GA1UdDwEB/wQEAwIHgDAbBgkrBgEEAYI3FQoEDjAMMAoGCCsGAQUFBwMDMB0G
# A1UdDgQWBBTEh+3Hi6JSfKfq9NrAKyDCC9JyLzAfBgNVHSMEGDAWgBSEq85L/ec2
# GZ2gu1TNr8SgdoHScDCCAVkGA1UdHwSCAVAwggFMMIIBSKCCAUSgggFAhoG7bGRh
# cDovLy9DTj1tZW1jb25maWdtZ3ItQ00xLUNBLENOPUNNMSxDTj1DRFAsQ049UHVi
# bGljJTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlv
# bixEQz1tZW1jb25maWdtZ3IsREM9b3JnP2NlcnRpZmljYXRlUmV2b2NhdGlvbkxp
# c3Q/YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0cmlidXRpb25Qb2ludIY+aHR0cDov
# L0NNMS5tZW1jb25maWdtZ3Iub3JnL0NlcnRFbnJvbGwvbWVtY29uZmlnbWdyLUNN
# MS1DQS5jcmyGQGZpbGU6Ly8vL0NNMS5tZW1jb25maWdtZ3Iub3JnL0NlcnRFbnJv
# bGwvbWVtY29uZmlnbWdyLUNNMS1DQS5jcmwwgcoGCCsGAQUFBwEBBIG9MIG6MIG3
# BggrBgEFBQcwAoaBqmxkYXA6Ly8vQ049bWVtY29uZmlnbWdyLUNNMS1DQSxDTj1B
# SUEsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29u
# ZmlndXJhdGlvbixEQz1tZW1jb25maWdtZ3IsREM9b3JnP2NBQ2VydGlmaWNhdGU/
# YmFzZT9vYmplY3RDbGFzcz1jZXJ0aWZpY2F0aW9uQXV0aG9yaXR5MDQGA1UdEQQt
# MCugKQYKKwYBBAGCNxQCA6AbDBlsYWJhZG1pbkBtZW1jb25maWdtZ3Iub3JnME0G
# CSsGAQQBgjcZAgRAMD6gPAYKKwYBBAGCNxkCAaAuBCxTLTEtNS0yMS0yNzU3MzIw
# NDAtOTUyNDUzNzU0LTI4ODk3NTA5NTAtMTEwMzANBgkqhkiG9w0BAQsFAAOCAQEA
# eoRuRgzEzrTSPnayr1mefYvXU8NRQcRUzZs4SYbs7cINqKK17EY5trDLrAR9FsNP
# Ya4ZcWa6nyaciK3iTZjeeqcHujmqA9viEMO/bpdak2yNfu60fwuH7OhHXcCYZf1n
# kCE7m2QuM5WJ2WWhNwx+lHKbvYs5z0DHweOUAe8kF5aVYjPXFbQEi7VFPhXJK7KV
# D/cFltTQgEwHowMysrnEYr3CHIf2HP4K2riL+5Wh5H0LX7af7HdRmo8jQnEHkYbH
# Zn+gilKGk6TXsiUVv5eqPPIGtTi6yA9m9dNbOslf9iK0oeISv2JJ0wGXSmOte+X7
# J4tH5/slqyFpky+23p0o0TGCAgkwggIFAgEBMGgwUTETMBEGCgmSJomT8ixkARkW
# A29yZzEcMBoGCgmSJomT8ixkARkWDG1lbWNvbmZpZ21ncjEcMBoGA1UEAxMTbWVt
# Y29uZmlnbWdyLUNNMS1DQQITJAAAAUPgt24PVLctvwAAAAABQzAJBgUrDgMCGgUA
# oHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYB
# BAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0B
# CQQxFgQU1vuZgPpmLub1tKx/xSHhMybmZ18wDQYJKoZIhvcNAQEBBQAEggEAQ8Ed
# Wg+d2lOexrESyUxETkiaJxEMFd5zesKYsismKmGON2UMpDyJZRISb7qj7/JHBJpv
# Dr2vK0JpD9RdKp8b7od5enMTV+OnY+jnyx8pnFAsSWCU03qWIdNG4NTQX/RH6oMF
# LMiyk+8EyOL4DJjL6SQIX6YVQlfgwKqAqKYVYwxssQFUV8mKocnLfyY/164A1ail
# 3YtU4P3y+6+XLDKI2ngbzYY4/uVojHNfg9IC1FLMnkJbBK59M7h8EyxRGxGBN08k
# 7WcNQSj27mC6co8EDvmgC+JJXuI31xNtgzlWa5ShoS2lRGXwmkrcY+IHCl6AFfPE
# py8aeKIUtAppd+Oimw==
# SIG # End signature block
