# Author: John Rea
# Date: May 1, 2024
# Version: 1.5

# Adds the System.Windows.Forms assembly.
Add-Type -AssemblyName System.Windows.Forms
# Clears the clipboard on launch
[System.Windows.Forms.Clipboard]::Clear()

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Prompt-O-Matic for Microsoft Designer'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(500, 530)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.Topmost = $true

# Create dropdown for Subject
$subjectLabel = New-Object System.Windows.Forms.Label
$subjectLabel.Text = 'Subject                       Freeze'
$subjectLabel.AutoSize = $true
$subjectLabel.Location = New-Object System.Drawing.Point(5, 10)
$subjectDropdown = New-Object System.Windows.Forms.ComboBox
$subjectDropdown.Location = New-Object System.Drawing.Point(5, 30)
$subjectDropdown.DropDownStyle = 'DropDown'
$subjectdropdown.items.addrange((@('an airplane', 'an alligator', 'an archipelago', 'a backyard', 'a balloon', 'a bay', 'a beach', 'a bear', 'a bedroom', 'a bicycle', 'a book', 'a boy', 'a bridge', 'a butterfly', 'a camera', 'a cape', 'a castle', 'a cat', 'a cave', 'a city', 'a cliff', 'a cloud', 'a computer', 'a coral', 'a crocodile', 'a den', 'a desert', 'a dessert', 'a dolphin', 'an eagle', 'an elephant', 'a flower', 'a forest', 'a fountain', 'a frog', 'a galaxy', 'a garden', 'a garage', 'a giraffe', 'a girl', 'a glacier', 'a guitar', 'a harbor', 'a helicopter', 'a hill', 'a hippopotamus', 'a horse', 'an iceberg', 'an island', 'a jellyfish', 'a kangaroo', 'a kitchen', 'a kite', 'a koala', 'a lamp', 'a lighthouse', 'a lion', 'a living room', 'a man', 'a meadow', 'a monkey', 'a moon', 'a mountain', 'a museum', 'an ocean', 'an octopus', 'an owl', 'a panda', 'a parrot', 'a patio', 'a peacock', 'a pebble', 'a penguin', 'a peninsula', 'a piano', 'a pond', 'a pyramid', 'a rabbit', 'a river', 'a rose', 'some sand', 'a seashell', 'a shark', 'a skyscraper', 'a snake', 'a squirrel', 'a star', 'a statue', 'a study', 'a sunset', 'a telescope', 'a temple', 'a tesla', 'a tiger', 'a train', 'a tree', 'a truck', 'a turtle', 'a pirate', 'an actor', 'a doctor', 'a dancer', 'a ranger', 'a valley', 'a volcano', 'a waterfall', 'a whale', 'a wolf', 'a woman', 'a zebra', 'a landscape', 'a portrait', 'a still life', 'an abstract', 'a cityscape', 'a seascape', 'some animals', 'some flowers', 'some trees', 'some rivers', 'some mountains', 'the sky', 'some clouds', 'a sunset', 'a sunrise', 'some buildings', 'some streets', 'some cars', 'some people', 'some children', 'some dancers', 'some musicians', 'some boats', 'some harbors', 'some beaches', 'some gardens', 'some forests', 'some deserts', 'some snow', 'some rain', 'some storms', 'some stars', 'the moon', 'some planets', 'some galaxies', 'some mythological creatures', 'some historical events', 'some cultural symbols', 'some religious icons', 'some angels', 'some demons', 'some heroes', 'some villains', 'some fantasy worlds', 'some science fiction scenes', 'some dreams', 'some nightmares', 'some love', 'some conflict', 'some peace', 'da vincis gaze', 'michelangelos creation', 'van goghs swirl', 'picassos fragment', 'monets bloom', 'dalis melt', 'rembrandts illuminate', 'matisses cut', 'pollocks drip', 'warhols replicate', 'hoppers isolate', 'klimts embrace', 'okeeffes enlarge', 'hokusais wave', 'banksys defy') | sort-object -unique))

# Create checkbox for each dropdown
$subjectCheckboxLabel = New-Object System.Windows.Forms.Label
$subjectCheckboxLabel.Location = New-Object System.Drawing.Point(130, 20)
$subjectCheckboxLabel.AutoSize = $true
$subjectCheckbox = New-Object System.Windows.Forms.CheckBox
$subjectCheckbox.Location = New-Object System.Drawing.Point(130, 30)
$subjectCheckbox.AutoSize = $true
$form.Controls.Add($subjectCheckbox)

# Create dropdown for Action
$actionLabel = New-Object System.Windows.Forms.Label
$actionLabel.Text = 'Action                       Freeze'
$actionLabel.AutoSize = $true
$actionLabel.Location = New-Object System.Drawing.Point(5, 60)
$actionDropdown = New-Object System.Windows.Forms.ComboBox
$actionDropdown.Location = New-Object System.Drawing.Point(5, 80)
$actionDropdown.DropDownStyle = 'DropDown'
$actiondropdown.items.addrange((@('acting', 'admiring', 'advertising', 'aiding', 'allowing', 'analyzing', 'appreciating', 'appointing', 'arguing', 'arranging', 'aspiring', 'assembling', 'assisting', 'auditing', 'authorizing', 'backing', 'bargaining', 'bartering', 'becoming', 'building', 'calculating', 'campaigning', 'capturing', 'caring', 'carving', 'celebrating', 'certifying', 'changing', 'charging', 'cherishing', 'choosing', 'climbing', 'coding', 'commanding', 'communicating', 'compensating', 'competing', 'completing', 'concluding', 'confirming', 'connecting', 'considering', 'constructing', 'contemplating', 'controlling', 'cooking', 'coordinating', 'creating', 'curing', 'cutting', 'cycling', 'dancing', 'debating', 'deciding', 'defending', 'defining', 'delivering', 'designing', 'detecting', 'developing', 'directing', 'discussing', 'discovering', 'distributing', 'diving', 'dock', 'drawing', 'dreaming', 'driving', 'editing', 'educating', 'electing', 'emancipating', 'employing', 'empowering', 'enabling', 'endorsing', 'enjoying', 'equipping', 'etching', 'examining', 'existing', 'explaining', 'exploring', 'expressing', 'extracting', 'feeling', 'filming', 'financing', 'fishing', 'flying', 'forging', 'forgetting', 'functioning', 'funding', 'governing', 'greeting', 'growing', 'guarding', 'guiding', 'hacking', 'harvesting', 'healing', 'hearing', 'helping', 'highlighting', 'hiring', 'holding', 'honoring', 'hoping', 'hunting', 'identifying', 'imagining', 'informing', 'innovating', 'inspecting', 'interpreting', 'investigating', 'jumping', 'keeping', 'labeling', 'landing', 'leading', 'leasing', 'learning', 'listening', 'living', 'loaning', 'lobbying', 'loving', 'maintaining', 'managing', 'manufacturing', 'marketing', 'measuring', 'mediating', 'medicating', 'mining', 'monitoring', 'molding', 'naming', 'navigating', 'negotiating', 'observing', 'offering', 'operating', 'orbiting', 'organizing', 'painting', 'performing', 'permitting', 'picking', 'planning', 'planting', 'playing', 'practicing', 'preparing', 'printing', 'processing', 'producing', 'programming', 'promoting', 'protecting', 'providing', 'publishing', 'racing', 'raising', 'reading', 'recalling', 'recording', 'recovering', 'refining', 'remembering', 'renting', 'rescuing', 'researching', 'resolving', 'respecting', 'rewarding', 'ruling', 'running', 'sailing', 'saving', 'sculpting', 'securing', 'seeing', 'selecting', 'selling', 'sensing', 'serving', 'shaping', 'shielding', 'shipping', 'singing', 'skating', 'skiing', 'smelling', 'solving', 'sponsoring', 'steering', 'studying', 'supporting', 'surfing', 'swimming', 'teaching', 'tracking', 'trading', 'trapping', 'traveling', 'treating', 'underlining', 'upholding', 'validating', 'viewing', 'voting', 'watching', 'winning', 'wishing', 'working', 'writing', 'depicting', 'capturing', 'rendering', 'illustrating', 'portraying', 'expressing', 'revealing', 'sketching', 'designing', 'crafting', 'shaping', 'forming', 'constructing', 'composing', 'creating', 'imagining', 'envisioning', 'dreaming', 'conceiving', 'cherishing', 'admiring', 'reflecting', 'meditating', 'contemplating', 'observing', 'viewing', 'seeing', 'discovering', 'uncovering', 'exposing', 'presenting', 'displaying', 'showcasing', 'featuring', 'highlighting', 'emphasizing', 'accentuating', 'illuminating', 'animating', 'energizing', 'vitalizing', 'revitalizing', 'awakening', 'stirring', 'arousing', 'inspiring', 'influencing', 'impacting', 'transforming', 'altering', 'changing') | sort-object -unique))

# Create checkbox for each dropdown
$actionCheckboxLabel = New-Object System.Windows.Forms.Label
$actionCheckboxLabel.Location = New-Object System.Drawing.Point(130, 80)
$actionCheckboxLabel.AutoSize = $true
$actionCheckbox = New-Object System.Windows.Forms.CheckBox
$actionCheckbox.Location = New-Object System.Drawing.Point(130, 90)
$actionCheckbox.AutoSize = $true
$form.Controls.Add($actionCheckbox)

# Create dropdown for Style
$styleLabel = New-Object System.Windows.Forms.Label
$styleLabel.Text = 'Style                       Freeze'
$styleLabel.AutoSize = $true
$styleLabel.Location = New-Object System.Drawing.Point(5, 110)
$styleDropdown = New-Object System.Windows.Forms.ComboBox
$styleDropdown.Location = New-Object System.Drawing.Point(5, 130)
$styleDropdown.DropDownStyle = 'DropDown'
$styledropdown.items.addrange((@('abstract art', 'academic art', 'acrylic paint', 'antique', 'arte povera', 'art brut', 'art deco', 'art nouveau', 'avant-garde', 'bauhaus', 'biological materials', 'bronze', 'canvas', 'caravaggism', 'ceramics', 'charcoal', 'classicism', 'clay', 'cobr', 'collage', 'color field painting', 'colored pencil', 'conceptual art', 'constructivism', 'contemporary art', 'cubism', 'cyber', 'dadaism', 'de stijl', 'digital art', 'encaustic', 'environmental materials', 'ephemeral materials', 'expressionism', 'fauvism', 'fiber', 'figurative art', 'fine art', 'found objects', 'futurism', 'futuristic', 'glass', 'gothic', 'graphite', 'gouache paint', 'harlem renaissance', 'ink', 'installation art', 'land art', 'leather', 'light', 'lyrical abstraction', 'marble', 'medieval', 'metal', 'minimalism', 'modern', 'modern art', 'mosaic', 'naive art', 'nautical', 'neo-dada', 'neo-expressionism', 'neo-impressionism', 'neoclassicism', 'neon art', 'oil paint', 'op art', 'paper', 'pastel', 'performance elements', 'photorealism', 'picassian', 'plastics', 'pop art', 'porcelain', 'post-impressionism', 'precisionism', 'realism', 'renaissance', 'rembrandt lighting', 'retro', 'rococo', 'romantic', 'sand', 'silverpoint', 'sound', 'steampunk', 'stone', 'street art', 'suprematism', 'surrealism', 'symbolism', 'tachisme', 'tempera paint', 'textiles', 'transavanguardia', 'video', 'vintage', 'watercolor paint', 'wood', 'zero group', 'surrealism', 'impressionism', 'realism', 'abstract expressionism', 'modernism', 'postmodernism', 'post-structuralism', 'cubism', 'fauvism', 'art nouveau', 'baroque', 'rococo', 'neoclassicism', 'romanticism', 'symbolism', 'expressionism', 'constructivism', 'dadaism', 'minimalism', 'pop art', 'op art', 'conceptual art', 'photorealism', 'installation art', 'performance art', 'digital art', 'street art', 'graffiti', 'land art', 'environmental art', 'bio art', 'video art', 'neo-expressionism', 'transavantgarde', 'neo-geo', 'young british artists', 'stuckism', 'superflat', 'relational aesthetics', 'vaporwave', 'lowbrow', 'toyism', 'post-internet', 'algorithmic art', 'generative art', 'virtual reality art', 'augmented reality art', 'cyber art', 'bio art', 'sound art', 'light art') | foreach-object {'in the ' + $_ + ' style'} | sort-object -unique))

# Create checkbox for each dropdown
$styleCheckboxLabel = New-Object System.Windows.Forms.Label
$styleCheckboxLabel.Location = New-Object System.Drawing.Point(130, 80)
$styleCheckboxLabel.AutoSize = $true
$styleCheckbox = New-Object System.Windows.Forms.CheckBox
$styleCheckbox.Location = New-Object System.Drawing.Point(130, 130)
$styleCheckbox.AutoSize = $true
$form.Controls.Add($styleCheckbox)

# Create dropdown for Media
$mediaLabel = New-Object System.Windows.Forms.Label
$mediaLabel.Text = 'Media                       Freeze'
$mediaLabel.AutoSize = $true
$mediaLabel.Location = New-Object System.Drawing.Point(5, 160)
$mediaDropdown = New-Object System.Windows.Forms.ComboBox
$mediaDropdown.Location = New-Object System.Drawing.Point(5, 180)
$mediaDropdown.DropDownStyle = 'DropDown'
$mediadropdown.items.addrange((@('acrylic paint', 'alabaster', 'bamboo', 'beeswax', 'biologicals', 'bone', 'brass', 'bronze', 'car parts', 'cement', 'ceramic', 'chalk', 'charcoal', 'clay', 'colored pencils', 'copper', 'coral', 'crayon', 'dammar', 'diamonds', 'digital art', 'dirt', 'dye', 'egg tempera', 'encaustic', 'environmental materials', 'ephemeral materials', 'feathers', 'felt', 'fiber', 'found objects', 'fresco', 'garbage', 'gesso', 'glass', 'gold', 'gold leaf', 'gouache paint', 'granite', 'graphite', 'graphite powder', 'gypsum', 'ink', 'ivory', 'jade', 'lacquer', 'latex', 'leather', 'light', 'limestone', 'linoleum', 'lithography', 'magnet', 'mahogany', 'marble', 'metal', 'mosaic', 'mylar', 'oak', 'obsidian', 'oil paint', 'papyrus', 'paper', 'pastel', 'pastel pencils', 'pewter', 'plaster', 'plastics', 'plexiglass', 'polymer clay', 'porcelain', 'porphyry', 'resin', 'rubber', 'sand', 'silk', 'silver', 'silverpoint', 'slate', 'soapstone', 'sound', 'steel', 'stone', 'tempera paint', 'textiles', 'terracotta', 'utensils', 'vellum', 'velvet', 'venetian plaster', 'video', 'wax', 'wood') | foreach-object {'using ' + $_ + ' as artistic medium'} | sort-object -unique))

# Create checkbox for each dropdown
$mediaCheckboxLabel = New-Object System.Windows.Forms.Label
$mediaCheckboxLabel.Location = New-Object System.Drawing.Point(130, 80)
$mediaCheckboxLabel.AutoSize = $true
$mediaCheckbox = New-Object System.Windows.Forms.CheckBox
$mediaCheckbox.Location = New-Object System.Drawing.Point(130, 180)
$mediaCheckbox.AutoSize = $true
$form.Controls.Add($mediaCheckbox)

# Create dropdown for Color
$colorLabel = New-Object System.Windows.Forms.Label
$colorLabel.Text = 'Color                       Freeze'
$colorLabel.AutoSize = $true
$colorLabel.Location = New-Object System.Drawing.Point(5, 210)
$colorDropdown = New-Object System.Windows.Forms.ComboBox
$colorDropdown.Location = New-Object System.Drawing.Point(5, 230)
$colorDropdown.DropDownStyle = 'DropDown'
$colordropdown.items.addrange((@('analogous', 'arctic', 'autumn', 'berry', 'black and white', 'blue', 'bright', 'candy', 'chocolate', 'citrus', 'coffee', 'complementary', 'cool', 'dark', 'desaturated', 'desert', 'dull', 'earth tones', 'fire', 'floral', 'forest', 'galactic', 'green', 'ice', 'jewel tones', 'light', 'metallic', 'midnight', 'monochrome', 'muted', 'nature', 'neon', 'neutral', 'oceanic', 'pastel', 'primary', 'psychedelic', 'rainbow', 'red', 'rose', 'rose gold', 'rustic', 'saturated', 'secondary', 'sepia', 'smoky', 'split-complementary', 'spring', 'stormy', 'summer', 'sunrise', 'sunset', 'tertiary', 'tetradic', 'triadic', 'tropical', 'twilight', 'urban', 'vibrant', 'warm', 'wine', 'winter', 'yellow') | foreach-object {'using a ' + $_ + ' color palette'} | sort-object -unique)) #>

# Create checkbox for each dropdown
$colorCheckboxLabel = New-Object System.Windows.Forms.Label
$colorCheckboxLabel.Location = New-Object System.Drawing.Point(130, 80)
$colorCheckboxLabel.AutoSize = $true
$colorCheckbox = New-Object System.Windows.Forms.CheckBox
$colorCheckbox.Location = New-Object System.Drawing.Point(130, 230)
$colorCheckbox.AutoSize = $true
$form.Controls.Add($colorCheckbox)

# Create dropdown for Extra A
$extraALabel = New-Object System.Windows.Forms.Label
$extraALabel.Text = 'ExtraA                       Freeze'
$extraALabel.AutoSize = $true
$extraALabel.Location = New-Object System.Drawing.Point(5, 260)
$extraADropdown = New-Object System.Windows.Forms.ComboBox
$extraADropdown.Location = New-Object System.Drawing.Point(5, 280)
$extraADropdown.DropDownStyle = 'DropDown'
$extraadropdown.items.addrange((@('8k', 'bokeh', 'crisp', 'double exposure', 'dramatic', 'dreamy', 'ethereal', 'extreme close-up', 'extremely detailed', 'glossy finish', 'golden hour', 'grainy', 'hazy', 'high contrast', 'high dynamic range', 'high key', 'infrared', 'kodachrome', 'landscape', 'leading lines', 'lens flare', 'long exposure', 'low key', 'macro', 'matte finish', 'minimalist', 'moody', 'motion blur', 'panoramic view', 'portrait', 'reflections', 'rule of thirds', 'shallow depth of field', 'silhouette', 'smoke filled', 'smooth', 'soft focus', 'star trails', 'studio lighting', 'symmetry', 'textured', 'tilt-shifted', 'time-lapse', 'upside down', 'vignette', 'vintage', 'whimsical') | sort-object -unique))

# Create checkbox for each dropdown
$extraACheckboxLabel = New-Object System.Windows.Forms.Label
$extraACheckboxLabel.Location = New-Object System.Drawing.Point(130, 80)
$extraACheckboxLabel.AutoSize = $true
$extraACheckbox = New-Object System.Windows.Forms.CheckBox
$extraACheckbox.Location = New-Object System.Drawing.Point(130, 280)
$extraACheckbox.AutoSize = $true
$form.Controls.Add($extraACheckbox)

# Create dropdown for Extra B
$extraBLabel = New-Object System.Windows.Forms.Label
$extraBLabel.Text = 'ExtraB                       Freeze'
$extraBLabel.AutoSize = $true
$extraBLabel.Location = New-Object System.Drawing.Point(5, 310)
$extraBDropdown = New-Object System.Windows.Forms.ComboBox
$extraBDropdown.Location = New-Object System.Drawing.Point(5, 330)
$extraBDropdown.DropDownStyle = 'DropDown'
$extrabdropdown.items.addrange((@('3d rendered', 'aerial view', 'ambient light', 'backlit', 'birds eye view', 'bokehlicious', 'chiaroscuro', 'cross processing', 'cyberpunk', 'deep depth of field', 'dystopian', 'fantasy', 'film grain', 'fish-eye lens', 'flat lay', 'hard shadows', 'isometric', 'lens flare', 'light painting', 'macro shot', 'magical realism', 'negative space', 'panoramic', 'perspective', 'post-apocalyptic', 'rim light', 'sci-fi', 'silhouette', 'soft shadows', 'split toning', 'steampunk', 'sunburst', 'telephoto', 'utopian', 'wide angle', 'worms eye view', 'zoomed in','abstract', 'allegorical', 'aperture', 'blue color temperature', 'blue hour', 'brightness', 'candlelight', 'conceptual', 'contrast', 'dawn', 'depressive tone', 'diurnal', 'dusk', 'figurative', 'futuristic', 'geometric', 'glow', 'golden hour', 'high aperture', 'high iso', 'high shutter speed', 'historical', 'illumination', 'long focal length', 'low exposure', 'luminance', 'medium saturation', 'medium white balance', 'moonlight', 'mythological', 'narrative', 'neon', 'nightfall', 'nocturnal', 'opposite hue', 'organic', 'out of focus', 'radiance', 'retro', 'seasonal', 'starlight', 'symbolic', 'timeless', 'twilight') | sort-object -unique))

# Create checkbox for each dropdown
$extraBCheckboxLabel = New-Object System.Windows.Forms.Label
$extraBCheckboxLabel.Location = New-Object System.Drawing.Point(130, 80)
$extraBCheckboxLabel.AutoSize = $true
$extraBCheckbox = New-Object System.Windows.Forms.CheckBox
$extraBCheckbox.Location = New-Object System.Drawing.Point(130, 330)
$extraBCheckbox.AutoSize = $true
$form.Controls.Add($extraBCheckbox)

# Create dropdown for Extra C
$extraCLabel = New-Object System.Windows.Forms.Label
$extraCLabel.Text = 'ExtraC                       Freeze'
$extraCLabel.AutoSize = $true
$extraCLabel.Location = New-Object System.Drawing.Point(5, 360)
$extraCDropdown = New-Object System.Windows.Forms.ComboBox
$extraCDropdown.Location = New-Object System.Drawing.Point(5, 380)
$extraCDropdown.DropDownStyle = 'DropDown'
$extracdropdown.items.addrange((@('aperture', 'shutter speed', 'iso', 'exposure', 'focal length', 'wide-angle lens', 'telephoto lens', 'prime lens', 'zoom lens', 'macro lens', 'fish-eye lens', 'depth of field', 'bokeh', 'resolution', 'pixel', 'sensor', 'mirrorless camera', 'dslr', 'point-and-shoot camera', 'medium format camera', 'large format camera', 'film camera', 'digital camera', 'compact camera', 'camera body', 'lens mount', 'image stabilization', 'tripod', 'monopod', 'flash', 'speedlight', 'reflector', 'softbox', 'light meter', 'raw format', 'jpeg format', 'white balance', 'histogram', 'exposure compensation', 'bracketing', 'time-lapse', 'long exposure', 'panorama', 'hdr', 'filter', 'polarizing filter', 'neutral density filter') | sort-object -unique))

# Create checkbox for each dropdown
$extraCCheckboxLabel = New-Object System.Windows.Forms.Label
$extraCCheckboxLabel.Location = New-Object System.Drawing.Point(130, 80)
$extraCCheckboxLabel.AutoSize = $true
$extraCCheckbox = New-Object System.Windows.Forms.CheckBox
$extraCCheckbox.Location = New-Object System.Drawing.Point(130, 380)
$extraCCheckbox.AutoSize = $true
$form.Controls.Add($extraCCheckbox)

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
$copiedTextbox.Size = New-Object System.Drawing.Size(475, 50)
$copiedTextbox.Multiline = $true
$copiedTextbox.ScrollBars = 'Vertical'

# Create a 'Random' button
$randomButton = New-Object System.Windows.Forms.Button
$randomButton.Text = 'Random Prompt Generator'
$randomButton.Location = New-Object System.Drawing.Point(150, 30)
$randomButton.Size = New-Object System.Drawing.Size(150, 180)

$randomButton.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show('Are you sure you want to generate a random prompt? This will overwrite any chosen dropdowns!', 'Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
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
        $clipboardText = [string]::Join(', ', $selectedWords)
        
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
$webMDButton.Location = New-Object System.Drawing.Point(150, 215)
$webMDButton.Size = New-Object System.Drawing.Size(150, 90)

# Disable the button initially
$webMDButton.Enabled = $copiedTextbox.Text -ne ''

$webMDButton.Add_Click({
    # Get the current clipboard text
    $clipboardText = Get-Clipboard

    if ($clipboardText) {
        $copiedTextbox.Text = $clipboardText
    } else {
        [System.Windows.Forms.MessageBox]::Show('Clipboard is empty. Please select at least one item.', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Define the base URL for the website
    $baseURL = 'https://designer.microsoft.com/image-creator?p='
    # Create a function to format the text for the URL
    function Format-ForURL ($text) {
        try {
            $text = $text.Replace(' ', '+')
            return $text
        } catch {
            [System.Windows.Forms.MessageBox]::Show('An error occurred while formatting the text: $_', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    # Open the website with the clipboard text
    $formattedText = Format-ForURL -text $copiedTextbox.Text
    if ($formattedText) {
        $url = $baseURL + $formattedText + '&form=NTPCH1&refig=dd8bc5ce95bf495a89a1b6c447914a00&pc=U531&adppc=EDGEDBB&sp=2&lq=0&qs=PN&sk=PN1&sc=8-0&cvid=dd8bc5ce95bf495a89a1b6c447914a00&showconv=1&sendquery=1'
        Start-Process $url
    } else {
        [System.Windows.Forms.MessageBox]::Show('Clipboard is empty. Please select at least one item.', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }

    # Log the clipboard text to a file
    if ($formattedText) {
        $logFile = 'prompt.log'
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = '$timestamp - $clipboardText'
        Add-Content -Path $logFile -Value $logEntry
    }
})

# Add an event handler to enable/disable the button based on the clipboard text
$copiedTextbox.add_TextChanged({
    $webMDButton.Enabled = $copiedTextbox.Text -ne ''
})

# Add a button to open the website with the generated text
$webMCButton = New-Object System.Windows.Forms.Button
$webMCButton.Text = 'Open Prompt in Microsoft Copilot'
$webMCButton.Location = New-Object System.Drawing.Point(150, 310)
$webMCButton.Size = New-Object System.Drawing.Size(150, 90)

# Disable the button initially
$webMCButton.Enabled = $copiedTextbox.Text -ne ''

$webMCButton.Add_Click({
    # Get the current clipboard text
    $clipboardText = Get-Clipboard

    if ($clipboardText) {
        $copiedTextbox.Text = $clipboardText
    } else {
        [System.Windows.Forms.MessageBox]::Show('Clipboard is empty. Please select at least one item.', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Define the base URL for the website
    $baseURL = 'https://www.bing.com/search?q='
    # Create a function to format the text for the URL
    function Format-ForURL ($text) {
        try {
            $text = $text.Replace(' ', '+')
            return $text
        } catch {
            [System.Windows.Forms.MessageBox]::Show('An error occurred while formatting the text: $_', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    # Open the website with the clipboard text
    $formattedText = Format-ForURL -text $copiedTextbox.Text
    if ($formattedText) {
        $url = $baseURL + 'an image of ' + $formattedText + '&form=NTPCH1&refig=dd8bc5ce95bf495a89a1b6c447914a00&pc=U531&adppc=EDGEDBB&sp=2&lq=0&qs=PN&sk=PN1&sc=8-0&cvid=dd8bc5ce95bf495a89a1b6c447914a00&showconv=1&sendquery=1&dissrchswrite=1'
        Start-Process $url
    } else {
        [System.Windows.Forms.MessageBox]::Show('Clipboard is empty. Please select at least one item.', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }

    # Log the clipboard text to a file
    if ($formattedText) {
        $logFile = 'prompt.log'
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = '$timestamp - $clipboardText'
        Add-Content -Path $logFile -Value $logEntry
    }
})

# Add an event handler to enable/disable the button based on the clipboard text
$copiedTextbox.add_TextChanged({
    $webMCButton.Enabled = $copiedTextbox.Text -ne ''
})

# Create a button to reset the dropdown boxes
$resetButton = New-Object System.Windows.Forms.Button
$resetButton.Location = New-Object System.Drawing.Point(305, 30)
$resetButton.Size = New-Object System.Drawing.Size(175, 370)
$resetButton.Text = 'Reset Dropdowns and Clear Clipboard'

# Add click event to the reset dropdown boxes button
$resetButton.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show('Are you sure you want to reset the dropdown boxes and clear the clipboard?', 'Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
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
$form.Controls.AddRange(@($subjectLabel, $subjectDropdown, $subjectLabelText, $actionLabel, $actionDropdown, $styleLabel, $styleDropdown, $mediaLabel, $mediaDropdown, $colorLabel, $colorDropdown, $extraALabel, $extraADropdown, $extraBLabel, $extraBDropdown,$extraCLabel, $extraCDropdown, $copybutton, $copiedLabel, $copiedTextbox))
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

# Add event handler to prompt for confirmation when closing the form
$form.Add_FormClosing({
    param($sender, $eventArgs)

    $confirmation = [System.Windows.Forms.MessageBox]::Show('Are you sure you want to close the form?', 'Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
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