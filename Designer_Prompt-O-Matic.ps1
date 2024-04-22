<#
Author: John Rea
Date: April 21, 2024
Version: 1.2
#>

#Adds the System.Windows.Forms assembly.
Add-Type -AssemblyName System.Windows.Forms

# Clear the clipboard on launch
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Clipboard]::Clear()

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
$copiedTextbox.Size = New-Object System.Drawing.Size(550, 100)

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
$subjectdropdown.items.addrange((@('cat', 'mountain', 'computer', 'ocean', 'painting', 'city', 'flower', 'book', 'river', 'bicycle', 'forest', 'museum', 'piano', 'island', 'balloon', 'guitar', 'castle', 'train', 'statue', 'garden', 'camera', 'bridge', 'tree', 'beach', 'telescope', 'skyscraper', 'airplane', 'star', 'pyramid', 'lamp', 'volcano', 'cloud', 'temple', 'truck', 'moon', 'waterfall', 'fountain', 'kite', 'sunset', 'helicopter', 'galaxy', 'cliff', 'lighthouse', 'rose', 'dolphin', 'elephant', 'lion', 'tiger', 'penguin', 'whale', 'shark', 'butterfly', 'zebra', 'horse', 'monkey', 'giraffe', 'bear', 'wolf', 'snake', 'eagle', 'parrot', 'owl', 'frog', 'turtle', 'rabbit', 'kangaroo', 'panda', 'alligator', 'crocodile', 'peacock', 'squirrel', 'koala', 'rhinoceros', 'hippopotamus', 'octopus', 'jellyfish', 'coral', 'seashell', 'sand', 'pebble', 'cave', 'iceberg', 'glacier', 'desert', 'canyon', 'plateau', 'valley', 'hill', 'meadow', 'pond', 'stream', 'harbor', 'bay', 'cape', 'peninsula', 'archipelago') | sort-object -Unique))

# Create dropdown for Action
$actionLabel = New-Object System.Windows.Forms.Label
$actionLabel.Text = 'Action Verb:'
$actionLabel.AutoSize = $true
$actionLabel.Location = New-Object System.Drawing.Point(10, 60)
$actionDropdown = New-Object System.Windows.Forms.ComboBox
$actionDropdown.Location = New-Object System.Drawing.Point(10, 80)
$actionDropdown.Items.AddRange((@('run', 'jump', 'swim', 'fly', 'read', 'write', 'sing', 'dance', 'climb', 'cook', 'paint', 'draw', 'travel', 'explore', 'discover', 'invent', 'create', 'build', 'design', 'calculate', 'analyze', 'measure', 'record', 'play', 'compete', 'win', 'lose', 'practice', 'perform', 'act', 'direct', 'produce', 'film', 'edit', 'broadcast', 'stream', 'launch', 'orbit', 'land', 'dock', 'charge', 'power', 'program', 'code', 'hack', 'encrypt', 'decrypt', 'navigate', 'steer', 'drive', 'sail', 'cycle', 'race', 'ski', 'skate', 'surf', 'dive', 'fish', 'hunt', 'track', 'trap', 'capture', 'rescue', 'save', 'heal', 'cure', 'treat', 'medicate', 'operate', 'investigate', 'solve', 'resolve', 'mediate', 'negotiate', 'argue', 'debate', 'discuss', 'converse', 'communicate', 'translate', 'interpret', 'express', 'describe', 'define', 'explain', 'inform', 'teach', 'learn', 'study', 'research', 'examine', 'inspect', 'observe', 'view', 'watch', 'see', 'hear', 'listen', 'taste', 'touch', 'feel', 'smell', 'sense', 'experience', 'enjoy', 'appreciate', 'cherish', 'love', 'admire', 'respect', 'honor', 'celebrate', 'commemorate', 'remember', 'forget', 'recall', 'imagine', 'dream', 'wish', 'hope', 'aspire', 'plan', 'prepare', 'organize', 'arrange', 'coordinate', 'manage', 'lead', 'guide', 'direct', 'control', 'command', 'govern', 'rule', 'decide', 'choose', 'select', 'pick', 'vote', 'elect', 'appoint', 'hire', 'employ', 'work', 'operate', 'function', 'perform', 'serve', 'assist', 'help', 'support', 'aid', 'care', 'nurture', 'foster', 'cultivate', 'grow', 'plant', 'harvest', 'mine', 'extract', 'refine', 'process', 'produce', 'manufacture', 'assemble', 'construct', 'build', 'engineer', 'develop', 'innovate', 'pioneer', 'trailblaze', 'lead', 'pave', 'forge', 'shape', 'mold', 'sculpt', 'carve', 'cut', 'etch', 'engrave', 'print', 'publish', 'release', 'distribute', 'market', 'sell', 'trade', 'barter', 'negotiate', 'bargain', 'deal', 'contract', 'lease', 'rent', 'loan', 'borrow', 'lend', 'invest', 'fund', 'finance', 'sponsor', 'endorse', 'promote', 'advertise', 'campaign', 'lobby', 'advocate', 'support', 'back', 'uphold', 'defend', 'protect', 'secure', 'safeguard', 'shield', 'guard', 'watch', 'monitor', 'supervise', 'oversee', 'inspect', 'audit', 'check', 'verify', 'confirm', 'validate', 'authenticate', 'certify', 'approve', 'authorize', 'permit', 'allow', 'enable', 'empower', 'equip', 'supply', 'provide', 'deliver', 'furnish', 'offer', 'present', 'gift', 'reward', 'compensate', 'pay', 'reimburse', 'refund', 'settle', 'resolve', 'clear', 'close', 'complete', 'finish', 'end', 'terminate', 'conclude', 'wrap', 'seal', 'sign', 'stamp', 'mark', 'label', 'tag', 'identify', 'name', 'call', 'title', 'caption', 'headline', 'feature', 'spotlight', 'highlight', 'emphasize', 'accentuate', 'stress', 'underline', 'underscore', 'bolster', 'strengthen', 'fortify', 'reinforce', 'support', 'back', 'uphold', 'maintain', 'preserve', 'keep', 'hold', 'retain', 'sustain', 'prolong', 'extend', 'continue', 'persist', 'persevere', 'endure', 'survive', 'live', 'exist', 'be', 'become', 'turn', 'transform', 'change', 'alter', 'modify', 'adjust', 'adapt', 'evolve', 'progress', 'advance', 'move', 'shift', 'transfer', 'transport', 'convey', 'carry', 'bring', 'take', 'fetch', 'deliver', 'send', 'dispatch', 'mail', 'post', 'ship', 'freight', 'haul', 'tow', 'drag', 'pull', 'push', 'lift', 'raise', 'elevate', 'hoist', 'suspend', 'hang', 'dangle', 'swing', 'rock', 'shake', 'jolt', 'jostle', 'bump', 'hit', 'strike', 'smack', 'slap', 'punch', 'kick', 'jab', 'stab', 'slash', 'slice', 'cut', 'chop', 'hack', 'saw', 'carve', 'whittle', 'shave', 'trim', 'clip', 'snip', 'cut', 'sever', 'split', 'divide', 'separate', 'part', 'detach', 'disconnect', 'disengage', 'release', 'free', 'liberate', 'emancipate', 'rescue', 'save', 'redeem', 'recover', 'retrieve', 'regain', 'reclaim', 'recapture', 'repossess', 'reacquire', 'reoccupy', 'reenter', 'return', 'revert', 'retrace', 'rewind', 'reverse', 'invert', 'flip', 'rotate', 'spin', 'twist', 'turn', 'pivot', 'swivel', 'tilt', 'lean', 'slant', 'angle', 'slope', 'incline', 'decline', 'descend', 'fall', 'drop', 'plunge', 'dive', 'sink', 'submerge', 'immerse', 'drown', 'flood', 'overflow', 'spill', 'leak', 'drip', 'trickle', 'stream', 'flow', 'pour', 'gush', 'erupt', 'explode', 'burst', 'blow', 'ignite', 'burn', 'flame', 'blaze', 'glow', 'shine', 'sparkle', 'twinkle', 'flicker', 'flash', 'beam', 'radiate', 'emit', 'discharge', 'release', 'exude', 'ooze', 'seep', 'weep', 'cry', 'sob', 'wail', 'moan', 'groan', 'sigh', 'whisper', 'murmur', 'mutter', 'grumble', 'complain', 'protest', 'object', 'oppose', 'resist', 'defy', 'challenge', 'confront', 'face', 'encounter', 'meet', 'greet', 'welcome', 'host', 'entertain', 'amuse', 'delight', 'please', 'satisfy', 'content', 'fulfill', 'accomplish', 'achieve', 'attain', 'reach', 'grasp', 'clutch', 'grip', 'hold', 'embrace', 'hug', 'kiss', 'caress', 'stroke', 'pat', 'rub', 'massage', 'knead', 'press', 'squeeze', 'crush', 'grind', 'pulverize', 'mash', 'smash', 'demolish', 'destroy', 'ruin', 'wreck', 'shatter', 'crack', 'fracture', 'break', 'snap', 'tear', 'rip', 'rend', 'shred', 'mangle', 'distort', 'warp', 'twist', 'bend', 'bow', 'arch', 'curve', 'coil', 'curl', 'loop', 'knot', 'tie', 'bind', 'fasten', 'secure', 'attach', 'affix', 'stick', 'glue', 'paste', 'tack', 'pin', 'nail', 'screw', 'bolt', 'rivet', 'weld', 'solder', 'fuse', 'join', 'unite', 'merge', 'combine', 'blend', 'mix', 'stir', 'whisk', 'beat', 'whip', 'froth', 'foam', 'bubble', 'fizz', 'effervesce', 'sparkle', 'pop', 'crackle', 'snap', 'hiss','sizzle', 'sear', 'scorch', 'char', 'blacken', 'toast', 'roast', 'grill', 'broil', 'bake', 'fry', 'saute') | Sort-Object -Unique))

# Create dropdown for Style
$styleLabel = New-Object System.Windows.Forms.Label
$styleLabel.Text = 'Artistic Style:'
$styleLabel.AutoSize = $true
$styleLabel.Location = New-Object System.Drawing.Point(10, 110)
$styleDropdown = New-Object System.Windows.Forms.ComboBox
$styleDropdown.Location = New-Object System.Drawing.Point(10, 130)
$styleDropdown.Items.AddRange((@('Oil Paint', 'Acrylic Paint', 'Watercolor Paint', 'Gouache Paint', 'Tempera Paint', 'Encaustic', 'Ink', 'Charcoal', 'Graphite', 'Pastel', 'Colored Pencil', 'Silverpoint', 'Clay', 'Metal', 'Wood', 'Stone', 'Marble', 'Bronze', 'Glass', 'Digital Art', 'Collage', 'Textiles', 'Plastics', 'Ceramics', 'Porcelain', 'Mosaic', 'Paper', 'Canvas', 'Leather', 'Fiber', 'Sand', 'Found Objects', 'Light', 'Sound', 'Video', 'Performance Elements', 'Ephemeral Materials', 'Environmental Materials', 'Biological Materials', 'Abstract Art', 'Abstract Expressionism', 'Academic Art', 'Art Deco', 'Art Nouveau', 'Avant-Garde', 'Baroque', 'Bauhaus', 'Classicism', 'CoBrA', 'Color Field Painting', 'Conceptual Art', 'Constructivism', 'Contemporary Art', 'Cubism', 'Dadaism', 'Digital Art', 'Expressionism', 'Fauvism', 'Figurative Art', 'Fine Art', 'Futurism', 'Gothic Art', 'Harlem Renaissance', 'Impressionism', 'Installation Art', 'Land Art', 'Minimalism', 'Modern Art', 'Naïve Art', 'Neo-Impressionism', 'Neoclassicism', 'Neon Art', 'Op Art', 'Photorealism', 'Pop Art', 'Post-Impressionism', 'Precisionism', 'Realism', 'Rococo', 'Street Art', 'Suprematism', 'Surrealism', 'Symbolism', 'Zero Group', 'Caravaggism', 'Rembrandt Lighting', 'Picassian', 'Renaissance', 'Baroque', 'Gothic','Impressionism', 'Expressionism', 'Cubism', 'Fauvism', 'Futurism', 'Dadaism', 'Surrealism', 'Abstract Expressionism', 'Minimalism', 'Pop Art', 'Conceptual Art', 'Op Art', 'Photorealism', 'Neo-Dada', 'Art Brut', 'Color Field Painting', 'Tachisme', 'Lyrical Abstraction', 'Arte Povera', 'Neo-Expressionism', 'Transavanguardia', 'Post-Impressionism', 'De Stijl', 'Bauhaus', 'Constructivism','nautical', 'vintage', 'retro', 'antique', 'modern', 'futuristic', 'cyber', 'steampunk', 'gothic', 'romantic', 'baroque', 'rococo', 'renaissance', 'medieval') | ForEach-Object {$_ + ' style'} | Sort-Object -Unique))

# Create dropdown for Media
$mediaLabel = New-Object System.Windows.Forms.Label
$mediaLabel.Text = 'Type of Media:'
$mediaLabel.AutoSize = $true
$mediaLabel.Location = New-Object System.Drawing.Point(10, 160)
$mediaDropdown = New-Object System.Windows.Forms.ComboBox
$mediaDropdown.Location = New-Object System.Drawing.Point(10, 180)
$mediaDropdown.Items.AddRange((@('Oil Paint', 'Acrylic Paint', 'Watercolor Paint', 'Gouache Paint', 'Tempera Paint', 'Encaustic', 'Ink', 'Charcoal', 'Graphite', 'Pastel', 'Colored Pencil', 'Silverpoint', 'Clay', 'Metal', 'Wood', 'Stone', 'Marble', 'Bronze', 'Glass', 'Digital Art', 'Collage', 'Textiles', 'Plastics', 'Ceramics', 'Porcelain', 'Mosaic', 'Paper', 'Canvas', 'Leather', 'Fiber', 'Sand', 'Found Objects', 'Light', 'Sound', 'Video', 'Performance Elements', 'Ephemeral Materials', 'Environmental Materials', 'Biological Materials','Alabaster', 'Bamboo', 'Beeswax', 'Bone', 'Brass', 'Cement', 'Chalk', 'Copper', 'Coral', 'Crayon', 'Dammar', 'Dye', 'Egg Tempera', 'Feathers', 'Felt', 'Fresco', 'Gesso', 'Gold Leaf', 'Granite', 'Graphite Powder', 'Gypsum', 'Ivory', 'Jade', 'Lacquer', 'Latex', 'Limestone', 'Linoleum', 'Lithography', 'Magnet', 'Mahogany', 'Mylar', 'Oak', 'Obsidian', 'Papyrus', 'Pastel Pencil', 'Pewter', 'Plaster', 'Plexiglass', 'Polymer Clay', 'Porphyry', 'Resin', 'Rubber', 'Silk', 'Slate', 'Soapstone', 'Steel', 'Terracotta', 'Vellum', 'Venetian Plaster', 'Wax') | ForEach-Object {$_ + ' as media'} | Sort-Object -Unique))

# Create dropdown for Color
$colorLabel = New-Object System.Windows.Forms.Label
$colorLabel.Text = 'Type of Color Palette:'
$colorLabel.AutoSize = $true
$colorLabel.Location = New-Object System.Drawing.Point(10, 210)
$colorDropdown = New-Object System.Windows.Forms.ComboBox
$colorDropdown.Location = New-Object System.Drawing.Point(10, 230)
$colorDropdown.Items.AddRange((@('autumn','winter','spring','summer','black and white','monochrome','sepia color','neon','primary','secondary','tertiary', 'complementary','analogous','triadic','split-complementary','tetradic','warm','cool','neutral','earth tones','jewel tones','dark','light', 'muted','saturated','desaturated','vibrant','dull','bright','light','dark', 'bright','muted','saturated','desaturated','vibrant','dull','pastel', 'rustic', 'metallic', 'psychedelic', 'oceanic', 'tropical', 'forest', 'urban', 'galactic', 'arctic', 'desert', 'sunset', 'sunrise', 'twilight', 'midnight', 'floral', 'citrus', 'berry', 'fire', 'ice', 'stormy', 'smoky', 'rainbow', 'candy', 'chocolate', 'coffee', 'wine', 'rose') | ForEach-Object {$_ + ' color palette'} | Sort-Object -Unique))

# Create dropdown for Extra A
$extraALabel = New-Object System.Windows.Forms.Label
$extraALabel.Text = 'Extra A:'
$extraALabel.AutoSize = $true
$extraALabel.Location = New-Object System.Drawing.Point(10, 260)
$extraADropdown = New-Object System.Windows.Forms.ComboBox
$extraADropdown.Location = New-Object System.Drawing.Point(10, 280)
$extraADropdown.Items.AddRange((@('8k', 'kodachrome', 'golden hour', 'extreme close-up', 'extremely detailed', 'studio lighting', 'lens flare', 'bokeh', 'high contrast', 'low key', 'high key', 'silhouette', 'motion blur', 'long exposure', 'shallow depth of field', 'rule of thirds', 'leading lines', 'symmetry', 'minimalist', 'vintage', 'grainy', 'moody', 'dreamy', 'ethereal', 'dramatic', 'whimsical', 'upside down', 'reflections', 'double exposure', 'panorama', 'tilt-shift', 'infrared', 'macro', 'time-lapse', 'star trails', 'smoke', 'portrait', 'landscape','soft focus', 'vibrant colors', 'sepia tone', 'saturated', 'desaturated', 'hazy', 'glossy finish', 'matte finish', 'textured', 'smooth', 'crisp', 'dynamic range', 'HDR', 'vignette', 'selective color', 'monochrome', 'duotone', 'triadic colors', 'complementary colors', 'analogous colors', 'split toning', 'cross processing', 'film grain', 'light painting', 'negative space', 'perspective', 'aerial view', 'birds eye view', 'worms eye view', 'panoramic', 'fish-eye lens', 'wide angle', 'telephoto', 'zoomed in', 'macro shot', 'depth of field', 'bokehlicious', 'sunburst', 'flare', 'backlit', 'rim light', 'ambient light', 'soft shadows', 'hard shadows', 'silhouette', 'chiaroscuro', 'flat lay', 'isometric', '3D render', 'hyperrealism', 'surreal', 'fantasy', 'sci-fi', 'post-apocalyptic', 'utopian', 'dystopian', 'steampunk', 'cyberpunk', 'magical realism', 'abstract', 'geometric', 'organic', 'figurative', 'narrative', 'conceptual', 'symbolic', 'allegorical', 'mythological', 'historical', 'retro', 'futuristic', 'timeless', 'seasonal', 'nocturnal', 'diurnal', 'golden hour', 'blue hour', 'twilight', 'dawn', 'dusk', 'nightfall', 'starlight', 'moonlight', 'candlelight', 'neon', 'glow', 'radiance', 'luminance', 'illumination', 'brightness', 'contrast', 'saturation', 'hue', 'tone', 'color temperature', 'white balance', 'exposure', 'shutter speed', 'aperture', 'ISO', 'focal length', 'focus') | Sort-Object -Unique))

# Create dropdown for Extra B
$extraBLabel = New-Object System.Windows.Forms.Label
$extraBLabel.Text = 'Extra B:'
$extraBLabel.AutoSize = $true
$extraBLabel.Location = New-Object System.Drawing.Point(10, 310)
$extraBDropdown = New-Object System.Windows.Forms.ComboBox
$extraBDropdown.Location = New-Object System.Drawing.Point(10, 330)
$extraBDropdown.Items.AddRange((@('8k', 'kodachrome', 'golden hour', 'extreme close-up', 'extremely detailed', 'studio lighting', 'lens flare', 'bokeh', 'high contrast', 'low key', 'high key', 'silhouette', 'motion blur', 'long exposure', 'shallow depth of field', 'rule of thirds', 'leading lines', 'symmetry', 'minimalist', 'vintage', 'grainy', 'moody', 'dreamy', 'ethereal', 'dramatic', 'whimsical', 'upside down', 'reflections', 'double exposure', 'panorama', 'tilt-shift', 'infrared', 'macro', 'time-lapse', 'star trails', 'smoke', 'portrait', 'landscape','soft focus', 'vibrant colors', 'sepia tone', 'saturated', 'desaturated', 'hazy', 'glossy finish', 'matte finish', 'textured', 'smooth', 'crisp', 'dynamic range', 'HDR', 'vignette', 'selective color', 'monochrome', 'duotone', 'triadic colors', 'complementary colors', 'analogous colors', 'split toning', 'cross processing', 'film grain', 'light painting', 'negative space', 'perspective', 'aerial view', 'birds eye view', 'worms eye view', 'panoramic', 'fish-eye lens', 'wide angle', 'telephoto', 'zoomed in', 'macro shot', 'depth of field', 'bokehlicious', 'sunburst', 'flare', 'backlit', 'rim light', 'ambient light', 'soft shadows', 'hard shadows', 'silhouette', 'chiaroscuro', 'flat lay', 'isometric', '3D render', 'hyperrealism', 'surreal', 'fantasy', 'sci-fi', 'post-apocalyptic', 'utopian', 'dystopian', 'steampunk', 'cyberpunk', 'magical realism', 'abstract', 'geometric', 'organic', 'figurative', 'narrative', 'conceptual', 'symbolic', 'allegorical', 'mythological', 'historical', 'retro', 'futuristic', 'timeless', 'seasonal', 'nocturnal', 'diurnal', 'golden hour', 'blue hour', 'twilight', 'dawn', 'dusk', 'nightfall', 'starlight', 'moonlight', 'candlelight', 'neon', 'glow', 'radiance', 'luminance', 'illumination', 'brightness', 'contrast', 'saturation', 'hue', 'tone', 'color temperature', 'white balance', 'exposure', 'shutter speed', 'aperture', 'ISO', 'focal length', 'focus') | Sort-Object -Unique))

# Create dropdown for Extra C
$extraCLabel = New-Object System.Windows.Forms.Label
$extraCLabel.Text = 'Extra C:'
$extraCLabel.AutoSize = $true
$extraCLabel.Location = New-Object System.Drawing.Point(10, 360)
$extraCDropdown = New-Object System.Windows.Forms.ComboBox
$extraCDropdown.Location = New-Object System.Drawing.Point(10, 380)
$extraCDropdown.Items.AddRange((@('8k', 'kodachrome', 'golden hour', 'extreme close-up', 'extremely detailed', 'studio lighting', 'lens flare', 'bokeh', 'high contrast', 'low key', 'high key', 'silhouette', 'motion blur', 'long exposure', 'shallow depth of field', 'rule of thirds', 'leading lines', 'symmetry', 'minimalist', 'vintage', 'grainy', 'moody', 'dreamy', 'ethereal', 'dramatic', 'whimsical', 'upside down', 'reflections', 'double exposure', 'panorama', 'tilt-shift', 'infrared', 'macro', 'time-lapse', 'star trails', 'smoke', 'portrait', 'landscape','soft focus', 'vibrant colors', 'sepia tone', 'saturated', 'desaturated', 'hazy', 'glossy finish', 'matte finish', 'textured', 'smooth', 'crisp', 'dynamic range', 'HDR', 'vignette', 'selective color', 'monochrome', 'duotone', 'triadic colors', 'complementary colors', 'analogous colors', 'split toning', 'cross processing', 'film grain', 'light painting', 'negative space', 'perspective', 'aerial view', 'birds eye view', 'worms eye view', 'panoramic', 'fish-eye lens', 'wide angle', 'telephoto', 'zoomed in', 'macro shot', 'depth of field', 'bokehlicious', 'sunburst', 'flare', 'backlit', 'rim light', 'ambient light', 'soft shadows', 'hard shadows', 'silhouette', 'chiaroscuro', 'flat lay', 'isometric', '3D render', 'hyperrealism', 'surreal', 'fantasy', 'sci-fi', 'post-apocalyptic', 'utopian', 'dystopian', 'steampunk', 'cyberpunk', 'magical realism', 'abstract', 'geometric', 'organic', 'figurative', 'narrative', 'conceptual', 'symbolic', 'allegorical', 'mythological', 'historical', 'retro', 'futuristic', 'timeless', 'seasonal', 'nocturnal', 'diurnal', 'golden hour', 'blue hour', 'twilight', 'dawn', 'dusk', 'nightfall', 'starlight', 'moonlight', 'candlelight', 'neon', 'glow', 'radiance', 'luminance', 'illumination', 'brightness', 'contrast', 'saturation', 'hue', 'tone', 'color temperature', 'white balance', 'exposure', 'shutter speed', 'aperture', 'ISO', 'focal length', 'focus') | Sort-Object -Unique))

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
$button.Text = 'Copy Dropdown Contents to Current Prompt'
$button.Location = New-Object System.Drawing.Point(140, 30)

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
    $clipboardText = $clipboardText.Substring(0,1).ToUpper()+$clipboardText.Substring(1) + "."
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
        $clipboardText = [string]::Join(', ', $selectedWords)
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
$webButton.Text = 'Open Current Prompt in Designer'
$webButton.Location = New-Object System.Drawing.Point(230, 220)
$webButton.Size = New-Object System.Drawing.Size(150, 180)
$webButton.Add_Click({

    # Get the current clipboard text
    $clipboardText = Get-Clipboard

    if ($clipboardText) {
        $copiedTextbox.Text = $clipboardText
    } else {
        [System.Windows.Forms.MessageBox]::Show("Clipboard is empty. Please select at least one item.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Open the website with the clipboard text
    $formattedText = Format-ForURL -text $copiedTextbox.Text
    $url = $baseURL + $formattedText
    Start-Process $url

    # Log the clipboard text to a file
    $logFile = "prompt.log"
    Add-Content -Path $logFile -Value $clipboardText

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

# Add click event to the clear button
$clearButton.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to clear the clipboard?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmation -eq "Yes") {
        try {
            if ([System.Windows.Forms.Clipboard]::ContainsText()) {
                [System.Windows.Forms.Clipboard]::Clear()
                $copiedTextbox.Text = ""
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred while clearing the clipboard: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
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
