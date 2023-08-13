from ursina import *
from ursina.prefabs.first_person_controller import FirstPersonController
#柏林噪声
from perlin_noise import PerlinNoise

app = Ursina()

grass_texture = load_texture('assets/grass_block.png')
stone_texture = load_texture('assets/stone_block.png')
brick_texture = load_texture('assets/brick_block.png')
dirt_texture = load_texture('assets/dirt_block.png')
sky_texture = load_texture('assets/skybox.png')
arm_texture = load_texture('assets/arm_texture.png')
punch_sound = Audio('assets/punch_sound',loop=False,autoplay=False)
block_pick = 1

window.fps_counter.enabled = False
window.exit_button.visible = False

scene.fog_color = color.white
scene.fog_density = 0.04

def input(key):
    if key == 'q' or key == 'escape':
        quit()

def update():
    global block_pick
    if held_keys['1']: block_pick = 1
    if held_keys['2']: block_pick = 2
    if held_keys['3']: block_pick = 3
    if held_keys['4']: block_pick = 4

    if held_keys['left mouse'] or held_keys['right mouse']:
        hand.active()
    else:
        hand.passive()

class Black(Button):
    def __init__(self,position=(0,0,0),texture = grass_texture):
        super().__init__(
            parent = scene,
            position = position,
            model = 'assets/block',
            origin_y = 0.5,
            texture = texture,
            color = color.color(0,0,random.uniform(0.9,1)),
            #highlight_color = color.green,
            scale = 0.5
        )

    def input(self,key):
        if self.hovered:
            if key == 'left mouse down':
                punch_sound.play()
                if block_pick == 1: black = Black(position = self.position+mouse.normal,texture=grass_texture)
                if block_pick == 2: black = Black(position = self.position+mouse.normal,texture=stone_texture)
                if block_pick == 3: black = Black(position = self.position+mouse.normal,texture=brick_texture)
                if block_pick == 4: black = Black(position = self.position+mouse.normal,texture=dirt_texture)
            if key == 'right mouse down':
                punch_sound.play()
                destroy(self)

class Sky(Entity):
    def __init__(self):
        super().__init__(
            parent = scene,
            model = 'sphere',
            texture = sky_texture,
            scale = 150,
            double_sided = True
        )

class Hand(Entity):
    def __init__(self):
        super().__init__(
            parent = camera.ui,
            model = 'assets/arm',
            texture = arm_texture,
            scale = 0.2,
            rotation = Vec3(150,-10,0),
            position = Vec2(0.4,-0.6)
        )
    
    def active(self):
        self.position = Vec2(0.3,-0.5)

    def passive(self):
        self.position = Vec2(0.4,-0.6)

noise = PerlinNoise(octaves = 3 ,seed = 2023)
scale = 24

for z in range(50):
    for x in range(50):
        block = Black(position = (x,0,z))
        block.y = floor(noise([x/scale,z/scale])*8)

player = FirstPersonController()
sky = Sky()
hand = Hand()

app.run()