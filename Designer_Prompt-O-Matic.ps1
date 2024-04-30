# Author: John Rea
# Date: April 29, 2024
# Version: 1.4

# Adds the System.Windows.Forms assembly.
Add-Type -AssemblyName System.Windows.Forms
# Clears the clipboard on launch
[System.Windows.Forms.Clipboard]::Clear()

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Prompt-O-Matic for Microsoft Designer'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(600, 530)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.Topmost = $true

# Create dropdown for Subject
$subjectLabel = New-Object System.Windows.Forms.Label
$subjectLabel.Text = 'Subject of Artwork:'
$subjectLabel.AutoSize = $true
$subjectLabel.Location = New-Object System.Drawing.Point(10, 10)
$subjectDropdown = New-Object System.Windows.Forms.ComboBox
$subjectDropdown.Location = New-Object System.Drawing.Point(10, 30)
$subjectdropdown.items.addrange((@('an airplane', 'an alligator', 'an archipelago', 'a backyard', 'a balloon', 'a bay', 'a beach', 'a bear', 'a bedroom', 'a bicycle', 'a book', 'a boy', 'a bridge', 'a butterfly', 'a camera', 'a cape', 'a castle', 'a cat', 'a cave', 'a city', 'a cliff', 'a cloud', 'a computer', 'a coral', 'a crocodile', 'a den', 'a desert', 'a dessert', 'a dolphin', 'an eagle', 'an elephant', 'a flower', 'a forest', 'a fountain', 'a frog', 'a galaxy', 'a garden', 'a garage', 'a giraffe', 'a girl', 'a glacier', 'a guitar', 'a harbor', 'a helicopter', 'a hill', 'a hippopotamus', 'a horse', 'an iceberg', 'an island', 'a jellyfish', 'a kangaroo', 'a kitchen', 'a kite', 'a koala', 'a lamp', 'a lighthouse', 'a lion', 'a living room', 'a man', 'a meadow', 'a monkey', 'a moon', 'a mountain', 'a museum', 'an ocean', 'an octopus', 'an owl', 'a panda', 'a parrot', 'a patio', 'a peacock', 'a pebble', 'a penguin', 'a peninsula', 'a piano', 'a pond', 'a pyramid', 'a rabbit', 'a river', 'a rose', 'some sand', 'a seashell', 'a shark', 'a skyscraper', 'a snake', 'a squirrel', 'a star', 'a statue', 'a study', 'a sunset', 'a telescope', 'a temple', 'a tesla', 'a tiger', 'a train', 'a tree', 'a truck', 'a turtle', 'a pirate', 'an actor', 'a doctor', 'a dancer', 'a ranger', 'a valley', 'a volcano', 'a waterfall', 'a whale', 'a wolf', 'a woman', 'a zebra', 'a landscape', 'a portrait', 'a still life', 'an abstract', 'a cityscape', 'a seascape', 'some animals', 'some flowers', 'some trees', 'some rivers', 'some mountains', 'the sky', 'some clouds', 'a sunset', 'a sunrise', 'some buildings', 'some streets', 'some cars', 'some people', 'some children', 'some dancers', 'some musicians', 'some boats', 'some harbors', 'some beaches', 'some gardens', 'some forests', 'some deserts', 'some snow', 'some rain', 'some storms', 'some stars', 'the moon', 'some planets', 'some galaxies', 'some mythological creatures', 'some historical events', 'some cultural symbols', 'some religious icons', 'some angels', 'some demons', 'some heroes', 'some villains', 'some fantasy worlds', 'some science fiction scenes', 'some dreams', 'some nightmares', 'some love', 'some conflict', 'some peace', 'da vincis gaze', 'michelangelos creation', 'van goghs swirl', 'picassos fragment', 'monets bloom', 'dalis melt', 'rembrandts illuminate', 'matisses cut', 'pollocks drip', 'warhols replicate', 'hoppers isolate', 'klimts embrace', 'okeeffes enlarge', 'hokusais wave', 'banksys defy') | sort-object -unique))
$subjectDropdown.AutoCompleteMode = 'suggestAppend'
$subjectDropdown.AutoCompleteSource = 'ListItems'
$subjectDropdown.DropDownStyle = 'DropDown'
$subjectDropdown.AutoCompleteCustomSource.AddRange($subjectDropdown.Items)
$subjectFilterTextbox = New-Object System.Windows.Forms.TextBox
$subjectFilterTextbox.Location = New-Object System.Drawing.Point(200, 10)

# Create dropdown for Action
$actionLabel = New-Object System.Windows.Forms.Label
$actionLabel.Text = 'Action Verb:'
$actionLabel.AutoSize = $true
$actionLabel.Location = New-Object System.Drawing.Point(10, 60)
$actionDropdown = New-Object System.Windows.Forms.ComboBox
$actionDropdown.Location = New-Object System.Drawing.Point(10, 80)
$actiondropdown.items.addrange((@('acting', 'admiring', 'advertising', 'aiding', 'allowing', 'analyzing', 'appreciating', 'appointing', 'arguing', 'arranging', 'aspiring', 'assembling', 'assisting', 'auditing', 'authorizing', 'backing', 'bargaining', 'bartering', 'becoming', 'building', 'calculating', 'campaigning', 'capturing', 'caring', 'carving', 'celebrating', 'certifying', 'changing', 'charging', 'cherishing', 'choosing', 'climbing', 'coding', 'commanding', 'communicating', 'compensating', 'competing', 'completing', 'concluding', 'confirming', 'connecting', 'considering', 'constructing', 'contemplating', 'controlling', 'cooking', 'coordinating', 'creating', 'curing', 'cutting', 'cycling', 'dancing', 'debating', 'deciding', 'defending', 'defining', 'delivering', 'designing', 'detecting', 'developing', 'directing', 'discussing', 'discovering', 'distributing', 'diving', 'dock', 'drawing', 'dreaming', 'driving', 'editing', 'educating', 'electing', 'emancipating', 'employing', 'empowering', 'enabling', 'endorsing', 'enjoying', 'equipping', 'etching', 'examining', 'existing', 'explaining', 'exploring', 'expressing', 'extracting', 'feeling', 'filming', 'financing', 'fishing', 'flying', 'forging', 'forgetting', 'functioning', 'funding', 'governing', 'greeting', 'growing', 'guarding', 'guiding', 'hacking', 'harvesting', 'healing', 'hearing', 'helping', 'highlighting', 'hiring', 'holding', 'honoring', 'hoping', 'hunting', 'identifying', 'imagining', 'informing', 'innovating', 'inspecting', 'interpreting', 'investigating', 'jumping', 'keeping', 'labeling', 'landing', 'leading', 'leasing', 'learning', 'listening', 'living', 'loaning', 'lobbying', 'loving', 'maintaining', 'managing', 'manufacturing', 'marketing', 'measuring', 'mediating', 'medicating', 'mining', 'monitoring', 'molding', 'naming', 'navigating', 'negotiating', 'observing', 'offering', 'operating', 'orbiting', 'organizing', 'painting', 'performing', 'permitting', 'picking', 'planning', 'planting', 'playing', 'practicing', 'preparing', 'printing', 'processing', 'producing', 'programming', 'promoting', 'protecting', 'providing', 'publishing', 'racing', 'raising', 'reading', 'recalling', 'recording', 'recovering', 'refining', 'remembering', 'renting', 'rescuing', 'researching', 'resolving', 'respecting', 'rewarding', 'ruling', 'running', 'sailing', 'saving', 'sculpting', 'securing', 'seeing', 'selecting', 'selling', 'sensing', 'serving', 'shaping', 'shielding', 'shipping', 'singing', 'skating', 'skiing', 'smelling', 'solving', 'sponsoring', 'steering', 'studying', 'supporting', 'surfing', 'swimming', 'teaching', 'tracking', 'trading', 'trapping', 'traveling', 'treating', 'underlining', 'upholding', 'validating', 'viewing', 'voting', 'watching', 'winning', 'wishing', 'working', 'writing', 'depicting', 'capturing', 'rendering', 'illustrating', 'portraying', 'expressing', 'revealing', 'sketching', 'designing', 'crafting', 'shaping', 'forming', 'constructing', 'composing', 'creating', 'imagining', 'envisioning', 'dreaming', 'conceiving', 'cherishing', 'admiring', 'reflecting', 'meditating', 'contemplating', 'observing', 'viewing', 'seeing', 'discovering', 'uncovering', 'exposing', 'presenting', 'displaying', 'showcasing', 'featuring', 'highlighting', 'emphasizing', 'accentuating', 'illuminating', 'animating', 'energizing', 'vitalizing', 'revitalizing', 'awakening', 'stirring', 'arousing', 'inspiring', 'influencing', 'impacting', 'transforming', 'altering', 'changing') | sort-object -unique))
$actionDropdown.AutoCompleteMode = 'suggestAppend'
$actionDropdown.AutoCompleteSource = 'ListItems'
$actionDropdown.DropDownStyle = 'DropDown'
$actionDropdown.AutoCompleteCustomSource.AddRange($actionDropdown.Items)
$actionFilterTextbox = New-Object System.Windows.Forms.TextBox
$actionFilterTextbox.Location = New-Object System.Drawing.Point(200, 30)

# Create dropdown for Style
$styleLabel = New-Object System.Windows.Forms.Label
$styleLabel.Text = 'Artistic Style:'
$styleLabel.AutoSize = $true
$styleLabel.Location = New-Object System.Drawing.Point(10, 110)
$styleDropdown = New-Object System.Windows.Forms.ComboBox
$styleDropdown.Location = New-Object System.Drawing.Point(10, 130)
$styledropdown.items.addrange((@('abstract art', 'academic art', 'acrylic paint', 'antique', 'arte povera', 'art brut', 'art deco', 'art nouveau', 'avant-garde', 'bauhaus', 'biological materials', 'bronze', 'canvas', 'caravaggism', 'ceramics', 'charcoal', 'classicism', 'clay', 'cobr', 'collage', 'color field painting', 'colored pencil', 'conceptual art', 'constructivism', 'contemporary art', 'cubism', 'cyber', 'dadaism', 'de stijl', 'digital art', 'encaustic', 'environmental materials', 'ephemeral materials', 'expressionism', 'fauvism', 'fiber', 'figurative art', 'fine art', 'found objects', 'futurism', 'futuristic', 'glass', 'gothic', 'graphite', 'gouache paint', 'harlem renaissance', 'ink', 'installation art', 'land art', 'leather', 'light', 'lyrical abstraction', 'marble', 'medieval', 'metal', 'minimalism', 'modern', 'modern art', 'mosaic', 'naive art', 'nautical', 'neo-dada', 'neo-expressionism', 'neo-impressionism', 'neoclassicism', 'neon art', 'oil paint', 'op art', 'paper', 'pastel', 'performance elements', 'photorealism', 'picassian', 'plastics', 'pop art', 'porcelain', 'post-impressionism', 'precisionism', 'realism', 'renaissance', 'rembrandt lighting', 'retro', 'rococo', 'romantic', 'sand', 'silverpoint', 'sound', 'steampunk', 'stone', 'street art', 'suprematism', 'surrealism', 'symbolism', 'tachisme', 'tempera paint', 'textiles', 'transavanguardia', 'video', 'vintage', 'watercolor paint', 'wood', 'zero group', 'surrealism', 'impressionism', 'realism', 'abstract expressionism', 'modernism', 'postmodernism', 'post-structuralism', 'cubism', 'fauvism', 'art nouveau', 'baroque', 'rococo', 'neoclassicism', 'romanticism', 'symbolism', 'expressionism', 'constructivism', 'dadaism', 'minimalism', 'pop art', 'op art', 'conceptual art', 'photorealism', 'installation art', 'performance art', 'digital art', 'street art', 'graffiti', 'land art', 'environmental art', 'bio art', 'video art', 'neo-expressionism', 'transavantgarde', 'neo-geo', 'young british artists', 'stuckism', 'superflat', 'relational aesthetics', 'vaporwave', 'lowbrow', 'toyism', 'post-internet', 'algorithmic art', 'generative art', 'virtual reality art', 'augmented reality art', 'cyber art', 'bio art', 'sound art', 'light art') | foreach-object {'in the ' + $_ + ' style'} | sort-object -unique))
$styleDropdown.AutoCompleteMode = 'suggestAppend'
$styleDropdown.AutoCompleteSource = 'ListItems'
$styleDropdown.DropDownStyle = 'DropDown'
$styleDropdown.AutoCompleteCustomSource.AddRange($styleDropdown.Items)
$styleFilterTextbox = New-Object System.Windows.Forms.TextBox
$styleFilterTextbox.Location = New-Object System.Drawing.Point(200, 50)

# Create dropdown for Media
$mediaLabel = New-Object System.Windows.Forms.Label
$mediaLabel.Text = 'Type of Media:'
$mediaLabel.AutoSize = $true
$mediaLabel.Location = New-Object System.Drawing.Point(10, 160)
$mediaDropdown = New-Object System.Windows.Forms.ComboBox
$mediaDropdown.Location = New-Object System.Drawing.Point(10, 180)
$mediadropdown.items.addrange((@('acrylic paint', 'alabaster', 'bamboo', 'beeswax', 'biologicals', 'bone', 'brass', 'bronze', 'car parts', 'cement', 'ceramic', 'chalk', 'charcoal', 'clay', 'colored pencils', 'copper', 'coral', 'crayon', 'dammar', 'diamonds', 'digital art', 'dirt', 'dye', 'egg tempera', 'encaustic', 'environmental materials', 'ephemeral materials', 'feathers', 'felt', 'fiber', 'found objects', 'fresco', 'garbage', 'gesso', 'glass', 'gold', 'gold leaf', 'gouache paint', 'granite', 'graphite', 'graphite powder', 'gypsum', 'ink', 'ivory', 'jade', 'lacquer', 'latex', 'leather', 'light', 'limestone', 'linoleum', 'lithography', 'magnet', 'mahogany', 'marble', 'metal', 'mosaic', 'mylar', 'oak', 'obsidian', 'oil paint', 'papyrus', 'paper', 'pastel', 'pastel pencils', 'pewter', 'plaster', 'plastics', 'plexiglass', 'polymer clay', 'porcelain', 'porphyry', 'resin', 'rubber', 'sand', 'silk', 'silver', 'silverpoint', 'slate', 'soapstone', 'sound', 'steel', 'stone', 'tempera paint', 'textiles', 'terracotta', 'utensils', 'vellum', 'velvet', 'venetian plaster', 'video', 'wax', 'wood') | foreach-object {'using ' + $_ + ' as artistic medium'} | sort-object -unique))
$mediaDropdown.AutoCompleteMode = 'suggestAppend'
$mediaDropdown.AutoCompleteSource = 'ListItems'
$mediaDropdown.DropDownStyle = 'DropDown'
$mediaDropdown.AutoCompleteCustomSource.AddRange($mediaDropdown.Items)
$mediaFilterTextbox = New-Object System.Windows.Forms.TextBox
$mediaFilterTextbox.Location = New-Object System.Drawing.Point(200, 70)

# Create dropdown for Color
$colorLabel = New-Object System.Windows.Forms.Label
$colorLabel.Text = 'Type of Color Palette:'
$colorLabel.AutoSize = $true
$colorLabel.Location = New-Object System.Drawing.Point(10, 210)
$colorDropdown = New-Object System.Windows.Forms.ComboBox
$colorDropdown.Location = New-Object System.Drawing.Point(10, 230)
$colordropdown.items.addrange((@('analogous', 'arctic', 'autumn', 'berry', 'black and white', 'blue', 'bright', 'candy', 'chocolate', 'citrus', 'coffee', 'complementary', 'cool', 'dark', 'desaturated', 'desert', 'dull', 'earth tones', 'fire', 'floral', 'forest', 'galactic', 'green', 'ice', 'jewel tones', 'light', 'metallic', 'midnight', 'monochrome', 'muted', 'nature', 'neon', 'neutral', 'oceanic', 'pastel', 'primary', 'psychedelic', 'rainbow', 'red', 'rose', 'rose gold', 'rustic', 'saturated', 'secondary', 'sepia', 'smoky', 'split-complementary', 'spring', 'stormy', 'summer', 'sunrise', 'sunset', 'tertiary', 'tetradic', 'triadic', 'tropical', 'twilight', 'urban', 'vibrant', 'warm', 'wine', 'winter', 'yellow') | foreach-object {'using a ' + $_ + ' color palette'} | sort-object -unique))
$colorDropdown.AutoCompleteMode = 'suggestAppend'
$colorDropdown.AutoCompleteSource = 'ListItems'
$colorDropdown.DropDownStyle = 'DropDown'
$colorDropdown.AutoCompleteCustomSource.AddRange($colorDropdown.Items)
$colorFilterTextbox = New-Object System.Windows.Forms.TextBox
$colorFilterTextbox.Location = New-Object System.Drawing.Point(200, 90)

# Create dropdown for Extra A
$extraALabel = New-Object System.Windows.Forms.Label
$extraALabel.Text = 'Extra A:'
$extraALabel.AutoSize = $true
$extraALabel.Location = New-Object System.Drawing.Point(10, 260)
$extraADropdown = New-Object System.Windows.Forms.ComboBox
$extraADropdown.Location = New-Object System.Drawing.Point(10, 280)
$extraadropdown.items.addrange((@('8k', 'bokeh', 'crisp', 'double exposure', 'dramatic', 'dreamy', 'ethereal', 'extreme close-up', 'extremely detailed', 'glossy finish', 'golden hour', 'grainy', 'hazy', 'high contrast', 'high dynamic range', 'high key', 'infrared', 'kodachrome', 'landscape', 'leading lines', 'lens flare', 'long exposure', 'low key', 'macro', 'matte finish', 'minimalist', 'moody', 'motion blur', 'panoramic view', 'portrait', 'reflections', 'rule of thirds', 'shallow depth of field', 'silhouette', 'smoke filled', 'smooth', 'soft focus', 'star trails', 'studio lighting', 'symmetry', 'textured', 'tilt-shifted', 'time-lapse', 'upside down', 'vignette', 'vintage', 'whimsical') | sort-object -unique))
$extraADropdown.AutoCompleteMode = 'suggestAppend'
$extraADropdown.AutoCompleteSource = 'ListItems'
$extraADropdown.DropDownStyle = 'DropDown'
$extraADropdown.AutoCompleteCustomSource.AddRange($extraADropdown.Items)
$extraAFilterTextbox = New-Object System.Windows.Forms.TextBox
$extraAFilterTextbox.Location = New-Object System.Drawing.Point(200, 110)

# Create dropdown for Extra B
$extraBLabel = New-Object System.Windows.Forms.Label
$extraBLabel.Text = 'Extra B:'
$extraBLabel.AutoSize = $true
$extraBLabel.Location = New-Object System.Drawing.Point(10, 310)
$extraBDropdown = New-Object System.Windows.Forms.ComboBox
$extraBDropdown.Location = New-Object System.Drawing.Point(10, 330)
$extrabdropdown.items.addrange((@('3d rendered', 'aerial view', 'ambient light', 'backlit', 'birds eye view', 'bokehlicious', 'chiaroscuro', 'cross processing', 'cyberpunk', 'deep depth of field', 'dystopian', 'fantasy', 'film grain', 'fish-eye lens', 'flat lay', 'hard shadows', 'isometric', 'lens flare', 'light painting', 'macro shot', 'magical realism', 'negative space', 'panoramic', 'perspective', 'post-apocalyptic', 'rim light', 'sci-fi', 'silhouette', 'soft shadows', 'split toning', 'steampunk', 'sunburst', 'telephoto', 'utopian', 'wide angle', 'worms eye view', 'zoomed in','abstract', 'allegorical', 'aperture', 'blue color temperature', 'blue hour', 'brightness', 'candlelight', 'conceptual', 'contrast', 'dawn', 'depressive tone', 'diurnal', 'dusk', 'figurative', 'futuristic', 'geometric', 'glow', 'golden hour', 'high aperture', 'high iso', 'high shutter speed', 'historical', 'illumination', 'long focal length', 'low exposure', 'luminance', 'medium saturation', 'medium white balance', 'moonlight', 'mythological', 'narrative', 'neon', 'nightfall', 'nocturnal', 'opposite hue', 'organic', 'out of focus', 'radiance', 'retro', 'seasonal', 'starlight', 'symbolic', 'timeless', 'twilight') | sort-object -unique))
$extraBDropdown.AutoCompleteMode = 'suggestAppend'
$extraBDropdown.AutoCompleteSource = 'ListItems'
$extraBDropdown.DropDownStyle = 'DropDown'
$extraBDropdown.AutoCompleteCustomSource.AddRange($extraBDropdown.Items)
$extraBFilterTextbox = New-Object System.Windows.Forms.TextBox
$extraBFilterTextbox.Location = New-Object System.Drawing.Point(200, 130)

# Create dropdown for Extra C
$extraCLabel = New-Object System.Windows.Forms.Label
$extraCLabel.Text = 'Extra C:'
$extraCLabel.AutoSize = $true
$extraCLabel.Location = New-Object System.Drawing.Point(10, 360)
$extraCDropdown = New-Object System.Windows.Forms.ComboBox
$extraCDropdown.Location = New-Object System.Drawing.Point(10, 380)
$extracdropdown.items.addrange((@('aperture', 'shutter speed', 'iso', 'exposure', 'focal length', 'wide-angle lens', 'telephoto lens', 'prime lens', 'zoom lens', 'macro lens', 'fish-eye lens', 'depth of field', 'bokeh', 'resolution', 'pixel', 'sensor', 'mirrorless camera', 'dslr', 'point-and-shoot camera', 'medium format camera', 'large format camera', 'film camera', 'digital camera', 'compact camera', 'camera body', 'lens mount', 'image stabilization', 'tripod', 'monopod', 'flash', 'speedlight', 'reflector', 'softbox', 'light meter', 'raw format', 'jpeg format', 'white balance', 'histogram', 'exposure compensation', 'bracketing', 'time-lapse', 'long exposure', 'panorama', 'hdr', 'filter', 'polarizing filter', 'neutral density filter') | sort-object -unique))
$extraCDropdown.AutoCompleteMode = 'suggestAppend'
$extraCDropdown.AutoCompleteSource = 'ListItems'
$extraCDropdown.DropDownStyle = 'DropDown'
$extraCDropdown.AutoCompleteCustomSource.AddRange($extraCDropdown.Items)
$extraCFilterTextbox = New-Object System.Windows.Forms.TextBox
$extraCFilterTextbox.Location = New-Object System.Drawing.Point(200, 150)

# Add event handlers to dropdowns
$subjectDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$actionDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$styleDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$mediaDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$colorDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$extraADropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$extraBDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)
$extraCDropdown.add_SelectedIndexChanged($dropdownSelectionChanged)

# Assuming $subjectDropdown, $actionDropdown, etc. are your dropdown objects
$dropdowns = @($subjectDropdown, $actionDropdown, $styleDropdown, $mediaDropdown, $colorDropdown, $extraADropdown, $extraBDropdown, $extraCDropdown)

foreach ($dropdown in $dropdowns) {
    $dropdown.Add_SelectedIndexChanged({
        $selectedWords = @()
        foreach ($dd in $dropdowns) {
            if ($dd.SelectedItem) { $selectedWords += $dd.SelectedItem }
        }
        $clipboardText = [string]::Join(', ', $selectedWords)
        Set-Clipboard -Value $clipboardText
        if ($clipboardText) {
            $copiedTextbox.Text = $clipboardText
        }
    })
}

# Create button to copy to clipboard
$copybutton = New-Object System.Windows.Forms.Button
$copybutton.Size = New-Object System.Drawing.Size(75, 370)
$copybutton.Text = 'Copy Dropdowns to the Clipboard'
$copybutton.Location = New-Object System.Drawing.Point(140, 30)
$copybutton.Enabled = $false

# Enable button only when all dropdowns have a selection
$subjectDropdown.add_SelectedIndexChanged({$copybutton.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem})
$actionDropdown.add_SelectedIndexChanged({$copybutton.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem})
$styleDropdown.add_SelectedIndexChanged({$copybutton.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem})
$mediaDropdown.add_SelectedIndexChanged({$copybutton.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem})
$colorDropdown.add_SelectedIndexChanged({$copybutton.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem})
$extraADropdown.add_SelectedIndexChanged({$copybutton.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem})
$extraBDropdown.add_SelectedIndexChanged({$copybutton.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem})
$extraCDropdown.add_SelectedIndexChanged({$copybutton.Enabled = $subjectDropdown.SelectedItem -and $actionDropdown.SelectedItem -and $styleDropdown.SelectedItem -and $mediaDropdown.SelectedItem -and $colorDropdown.SelectedItem -and $extraADropdown.SelectedItem -and $extraBDropdown.SelectedItem -and $extraCDropdown.SelectedItem})
$copybutton.Add_Click({
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
        [System.Windows.Forms.MessageBox]::Show('Please select at least one item before clicking the button.', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $clipboardText = [string]::Join(', ', $selectedWords)
    Set-Clipboard -Value $clipboardText
    if ($clipboardText) {
        $copiedTextbox.Text = $clipboardText
    }
})

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

# Create a 'Random' button
$randomButton = New-Object System.Windows.Forms.Button
$randomButton.Text = 'Random Prompt Generator'
$randomButton.Location = New-Object System.Drawing.Point(230, 30)
$randomButton.Size = New-Object System.Drawing.Size(150, 180)
$randomButton.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show('Are you sure you want to generate a random prompt? This will overwrite any chosen dropdowns!', 'Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmation -eq 'Yes') {
        $selectedWords = @()
        $selectedWords += $subjectDropdown.Items | Get-Random
        $selectedWords += $actionDropdown.Items | Get-Random
        $selectedWords += $styleDropdown.Items | Get-Random
        $selectedWords += $mediaDropdown.Items | Get-Random
        $selectedWords += $colorDropdown.Items | Get-Random
        $selectedWords += $extraADropdown.Items | Get-Random
        $selectedWords += $extraBDropdown.Items | Get-Random
        $selectedWords += $extraCDropdown.Items | Get-Random
        $clipboardText = [string]::Join(', ', $selectedWords)
        if ($clipboardText) {
            #$clipboardText = $clipboardText.Substring(0,1).ToUpper()+$clipboardText.Substring(1)
            try {
                Set-Clipboard -Value $clipboardText
            } catch {
                [System.Windows.Forms.MessageBox]::Show('An error occurred while setting the clipboard: $_', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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

# Add the selected items from the Dropdowns to the selectedWords array
if ($subjectDropdown.Text) { $selectedWords += $subjectDropdown.Text }
if ($actionDropdown.Text) { $selectedWords += $actionDropdown.Text }
if ($styleDropdown.Text) { $selectedWords += $styleDropdown.Text }
if ($mediaDropdown.Text) { $selectedWords += $mediaDropdown.Text }
if ($colorDropdown.Text) { $selectedWords += $colorDropdown.Text }
if ($extraADropdown.Text) { $selectedWords += $extraADropdown.Text }
if ($extraBDropdown.Text) { $selectedWords += $extraBDropdown.Text }
if ($extraCDropdown.Text) { $selectedWords += $extraCDropdown.Text }

# Set the copied text to the joined selected words
if ($selectedWords -ne $null) {
    $copiedTextbox.Text = $selectedWords -join ' '
} else {
    if ($copiedTextbox -ne $null) {
        $copiedTextbox.Text = ''
    }
}

# Add a button to open the website with the generated text
$webMDButton = New-Object System.Windows.Forms.Button
$webMDButton.Text = 'Open Prompt in Microsoft Designer'
$webMDButton.Location = New-Object System.Drawing.Point(230, 220)
$webMDButton.Size = New-Object System.Drawing.Size(150, 85)

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
$webMCButton.Location = New-Object System.Drawing.Point(230, 315)
$webMCButton.Size = New-Object System.Drawing.Size(150, 85)

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
$resetButton.Location = New-Object System.Drawing.Point(400, 30)
$resetButton.Size = New-Object System.Drawing.Size(150, 180)
$resetButton.Text = 'Reset Dropdowns'

# Add click event to the reset dropdown boxes button
$resetButton.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show('Are you sure you want to reset the dropdown boxes?', 'Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmation -eq 'Yes') {
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
$clearButton.Text = 'Clear Clipboard'
# Disable the button initially
$clearButton.Enabled = [System.Windows.Forms.Clipboard]::ContainsText()

# Add click event to the clear button
$clearButton.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show('Are you sure you want to clear the clipboard?', 'Confirmation', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmation -eq 'Yes') {
        try {
            if ([System.Windows.Forms.Clipboard]::ContainsText()) {
                [System.Windows.Forms.Clipboard]::Clear()
                $copiedTextbox.Text = ''
                $clearButton.Enabled = $false
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show('An error occurred while clearing the clipboard: $_', 'Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Add event handler to enable/disable the clear button based on clipboard content
$copiedTextbox.add_TextChanged({
    $clearButton.Enabled = [System.Windows.Forms.Clipboard]::ContainsText()
})

# Add event handler to enable/disable the reset button based on dropdown selection
$subjectDropdown.add_SelectedIndexChanged({$resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem})
$actionDropdown.add_SelectedIndexChanged({$resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem})
$styleDropdown.add_SelectedIndexChanged({$resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem})
$mediaDropdown.add_SelectedIndexChanged({$resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem})
$colorDropdown.add_SelectedIndexChanged({$resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem})
$extraADropdown.add_SelectedIndexChanged({$resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem})
$extraBDropdown.add_SelectedIndexChanged({$resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem})
$extraCDropdown.add_SelectedIndexChanged({$resetButton.Enabled = $subjectDropdown.SelectedItem -or $actionDropdown.SelectedItem -or $styleDropdown.SelectedItem -or $mediaDropdown.SelectedItem -or $colorDropdown.SelectedItem -or $extraADropdown.SelectedItem -or $extraBDropdown.SelectedItem -or $extraCDropdown.SelectedItem})

# Add controls to the form
$form.Controls.Add($webMDButton)
$form.Controls.Add($webMCButton)
$form.Controls.Add($randomButton)
$form.Controls.Add($copiedLabel)
$form.Controls.Add($clearButton)
$form.Controls.Add($copybutton)
$form.Controls.Add($resetButton)
$form.Controls.AddRange(@($subjectLabel, $subjectDropdown, $actionLabel, $actionDropdown, $styleLabel, $styleDropdown, $mediaLabel, $mediaDropdown, $colorLabel, $colorDropdown, $extraALabel, $extraADropdown, $extraBLabel, $extraBDropdown,$extraCLabel, $extraCDropdown, $copybutton, $copiedLabel, $copiedTextbox))
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'

# Add event handler to prompt for confirmation when closing the form
$form.Add_FormClosing({
    param($sender, $eventArgs)  # Add this line

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