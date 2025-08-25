Attribute VB_Name = "ModMain"

Option Explicit

Public Const itp_type_horse = &H1
Public Const itp_type_one_handed_wpn = &H2
Public Const itp_type_two_handed_wpn = &H3
Public Const itp_type_polearm = &H4
Public Const itp_type_arrows = &H5
Public Const itp_type_bolts = &H6
Public Const itp_type_shield = &H7
Public Const itp_type_bow = &H8
Public Const itp_type_crossbow = &H9
Public Const itp_type_thrown = &HA
Public Const itp_type_goods = &HB
Public Const itp_type_head_armor = &HC
Public Const itp_type_body_armor = &HD
Public Const itp_type_foot_armor = &HE
Public Const itp_type_hand_armor = &HF
Public Const itp_type_pistol = &H10
Public Const itp_type_musket = &H11
Public Const itp_type_bullets = &H12
Public Const itp_type_animal = &H13
Public Const itp_type_book = &H14
Public Const itp_type_mask = &H1F

Public Const itp_force_attach_left_hand = "0000000000000100"
Public Const itp_force_attach_right_hand = "0000000000000200"
Public Const itp_force_attach_left_forearm = "0000000000000300"
Public Const itp_attach_armature = "0000000000000f00"
Public Const itp_attachment_mask = "0000000000000f00"
Public Const itp_Attachment_Left_bit = 8
Public Const itp_Attachment_Right_bit = 9
Public Const itp_Attachment_Armature_bit1 = 10
Public Const itp_Attachment_Armature_bit2 = 11

Public Const itp_unique = 12
Public Const itp_always_loot = 13
Public Const itp_no_parry = 14
Public Const itp_default_ammo = 15
Public Const itp_merchandise = 16
Public Const itp_wooden_attack = 17
Public Const itp_wooden_parry = 18
Public Const itp_food = 19

Public Const itp_cant_reload_on_horseback = 20
Public Const itp_two_handed = 21
Public Const itp_primary = 22
Public Const itp_secondary = 23
Public Const itp_covers_legs = 24
Public Const itp_doesnt_cover_hair = 24
Public Const itp_can_penetrate_shield = 24
Public Const itp_consumable = 25
Public Const itp_bonus_against_shield = 26
Public Const itp_penalty_with_shield = 27
Public Const itp_cant_use_on_horseback = 28
Public Const itp_civilian = 29
Public Const itp_next_item_as_melee = 29
Public Const itp_fit_to_head = 30
Public Const itp_offset_lance = 30
Public Const itp_covers_head = 31
Public Const itp_couchable = 31
Public Const itp_crush_through = 32
Public Const itp_knock_back = 33
Public Const itp_remove_item_on_use = 34
Public Const itp_unbalanced = 35

Public Const itp_covers_beard = 36
Public Const itp_no_pick_up_from_ground = 37
Public Const itp_can_knock_down = 38

Public Const itcf_Thrust_onehanded = "0000000000000001"
Public Const itcf_Overswing_onehanded = "0000000000000002"
Public Const itcf_Slashright_onehanded = "0000000000000004"
Public Const itcf_Slashleft_onehanded = "0000000000000008"

Public Const itcf_Thrust_twohanded = "0000000000000010"
Public Const itcf_Overswing_twohanded = "0000000000000020"
Public Const itcf_Slashright_twohanded = "0000000000000040"
Public Const itcf_Slashleft_twohanded = "0000000000000080"

Public Const itcf_Thrust_polearm = "0000000000000100"
Public Const itcf_Overswing_polearm = "0000000000000200"
Public Const itcf_Slashright_polearm = "0000000000000400"
Public Const itcf_Slashleft_polearm = "0000000000000800"

Public Const itcf_Shoot_bow = "0000000000001000"
Public Const itcf_Shoot_javelin = "0000000000002000"
Public Const itcf_Shoot_crossbow = "0000000000004000"

Public Const itcf_Throw_stone = "0000000000010000"
Public Const itcf_Throw_knife = "0000000000020000"
Public Const itcf_Throw_axe = "0000000000030000"
Public Const itcf_Throw_javelin = "0000000000040000"
Public Const itcf_Shoot_pistol = "0000000000070000"
Public Const itcf_Shoot_musket = "0000000000080000"
Public Const itcf_Shoot_mask = "00000000000ff000"

Public Const itcf_Horseback_thrust_onehanded = "0000000000100000"
Public Const itcf_Horseback_overswing_right_onehanded = "0000000000200000"
Public Const itcf_Horseback_overswing_left_onehanded = "0000000000400000"
Public Const itcf_Horseback_slashright_onehanded = "0000000000800000"
Public Const itcf_Horseback_slashleft_onehanded = "0000000001000000"
Public Const itcf_Thrust_onehanded_lance = "0000000004000000"
Public Const itcf_Thrust_onehanded_lance_horseback = "0000000008000000"

Public Const itcf_Carry_mask = "00000007f0000000"
Public Const itcf_Carry_sword_left_hip = "0000000010000000"
Public Const itcf_Carry_axe_left_hip = "0000000020000000"
Public Const itcf_Carry_dagger_front_left = "0000000030000000"
Public Const itcf_Carry_dagger_front_right = "0000000040000000"
Public Const itcf_Carry_quiver_front_right = "0000000050000000"
Public Const itcf_Carry_quiver_back_right = "0000000060000000"
Public Const itcf_Carry_quiver_right_vertical = "0000000070000000"
Public Const itcf_Carry_quiver_back = "0000000080000000"
Public Const itcf_Carry_revolver_right = "0000000090000000"
Public Const itcf_Carry_pistol_front_left = "00000000a0000000"
Public Const itcf_Carry_bowcase_left = "00000000b0000000"
Public Const itcf_Carry_mace_left_hip = "00000000c0000000"

Public Const itcf_Carry_axe_back = "0000000100000000"
Public Const itcf_Carry_sword_back = "0000000110000000"
Public Const itcf_Carry_kite_shield = "0000000120000000"
Public Const itcf_Carry_round_shield = "0000000130000000"
Public Const itcf_Carry_buckler_left = "0000000140000000"
Public Const itcf_Carry_crossbow_back = "0000000150000000"
Public Const itcf_Carry_bow_back = "0000000160000000"
Public Const itcf_Carry_spear = "0000000170000000"
Public Const itcf_Carry_board_shield = "0000000180000000"

Public Const itcf_Carry_katana = "0000000210000000"
Public Const itcf_Carry_wakizashi = "0000000220000000"


Public Const itcf_Show_holster_when_drawn = "0000000800000000"

Public Const itcf_Reload_pistol = "0000007000000000"
Public Const itcf_Reload_musket = "0000008000000000"
Public Const itcf_Reload_mask = "000000f000000000"

Public Const itcf_Parry_forward_onehanded = "0000010000000000"
Public Const itcf_Parry_up_onehanded = "0000020000000000"
Public Const itcf_Parry_right_onehanded = "0000040000000000"
Public Const itcf_Parry_left_onehanded = "0000080000000000"

Public Const itcf_Parry_forward_twohanded = "0000100000000000"
Public Const itcf_Parry_up_twohanded = "0000200000000000"
Public Const itcf_Parry_right_twohanded = "0000400000000000"
Public Const itcf_Parry_left_twohanded = "0000800000000000"

Public Const itcf_Parry_forward_polearm = "0001000000000000"
Public Const itcf_Parry_up_polearm = "0002000000000000"
Public Const itcf_Parry_right_polearm = "0004000000000000"
Public Const itcf_Parry_left_polearm = "0008000000000000"

Public Const itcf_Horseback_slash_polearm = "0010000000000000"

Public Const itcf_Force_64_bits = "8000000000000000"

Public Const imodbits_polearm = "8202"
Public Const imodbits_axe = "262164"
Public Const imodbits_missile = "4398046511112"
Public Const imodbits_bow = "655370"
Public Const imodbits_crossbow = "131082"
Public Const imodbits_horse_basic = "41876193280"
Public Const imodbits_shield = "167772194"
Public Const imodbits_sword = "24596"
Public Const imodbits_cloth = "123731968"
Public Const imodbits_plate = "704643238"
Public Const imodbits_horse_good = "110595670016"
Public Const imodbits_sword_high = "155668"
Public Const imodbits_pick = "270356"
Public Const imodbits_mace = "262164"
Public Const imodbits_thrown = "4398046781448"
Public Const imodbits_thrown_minus_heavy = "4398046519304"
Public Const imodbits_armor = "704643236"

Public Const ixmesh_inventory = "1000000000000000"
Public Const ixmesh_flying_ammo = "2000000000000000"
Public Const ixmesh_Carry = "3000000000000000"
Public Const ixmesh_Inventory_bit = 60
Public Const ixmesh_Flying_Ammo_bit = 61

Public Const ti_on_init_item = -50           'can only be used in module_items triggers
Public Const ti_on_weapon_attack = -51       'can only be used in module_items triggers
' Position Register 1: Weapon Item Position
Public Const ti_on_missile_hit = -52         'can only be used in module_items triggers
' Position Register 1: Missile Position
' Trigger Param 1: shooter agent id
Public Const ti_on_init_map_icon = -70                   'can only be used in module_map_icons triggers
' Trigger Param 1: id of the owner party

'����ϵͳ
Public Const psf_always_emit = "0000000002"
Public Const psf_global_emit_dir = "0000000010"
Public Const psf_emit_at_water_level = "0000000020"
Public Const psf_billboard_2d = "0000000100"                  '# up_vec = dir, front rotated towards camera
Public Const psf_billboard_3d = "0000000200"                  '  # front_vec point to camera.
Public Const psf_billboard_drop = "0000000300"
Public Const psf_turn_to_velocity = "0000000400"
Public Const psf_randomize_rotation = "0000001000"
Public Const psf_randomize_size = "0000002000"
Public Const psf_2d_turbulance = "0000010000"

'����ģ��
Public Const render_order_plus_1 = "00000001"

'Tags
Public Const Tag_Register = 1
Public Const Tag_Variable = 2
Public Const Tag_String = 3
Public Const Tag_Item = 4
Public Const Tag_Troop = 5
Public Const Tag_Faction = 6
Public Const Tag_Quest = 7
Public Const Tag_Party_Tpl = 8
Public Const Tag_Party = 9
Public Const Tag_Scene = 10
Public Const Tag_Mission_tpl = 11
Public Const Tag_Menu = 12
Public Const Tag_Script = 13
Public Const Tag_Particle_Sys = 14
Public Const Tag_Scene_Prop = 15
Public Const Tag_Sound = 16
Public Const Tag_Local_Variable = 17
Public Const Tag_Map_Icon = 18
Public Const Tag_Skill = 19
Public Const Tag_Mesh = 20
Public Const Tag_Presentation = 21
Public Const Tag_Quick_String = 22
Public Const Tag_Track = 23
Public Const Tag_Tableau = 24
Public Const Tag_Animation = 25
Public Const Tags_End = 26

Public Tags(1 To 26) As String

Public Const N_Pos = 65

Public Const Max_Indentation = 16

Type type_Attrib
    '#  9) Attributes (int): Example usage:
    '#    str_6|agi_6|int_4|cha_5|level(5)
    strPoint As Integer
    agiPoint As Integer
    intPoint As Integer
    chaPoint As Integer
    level As Integer
End Type

Type Type_WeaponProf
    '# 10) Weapon proficiencies (int): Example usage:
    '# wp_one_handed(55)|wp_two_handed(90)|wp_polearm(36)|wp_archery(80)|wp_crossbow(24)|wp_throwing(45)
    one_handed As Long
    two_handed As Long
    polearm As Long
    archery As Long
    crossbow As Long
    throwing As Long
    firearm As Long
End Type

Type Type_XY
    X As Long
    Y As Long
End Type

Type Type_XY_Index
    X As Long
    strX As String
    Y As Long
End Type

Type Type_strXY
    X As String
    Y As String
End Type

Type Type_strXYZ
    X As String
    Y As String
    Z As String
End Type

Type Type_dblXYZ
    X As Double
    Y As Double
    Z As Double
End Type

Public Type Type_RelationShip
    ID As Long
    strID As String
    Value As Double
End Type

Public Type Type_Chest
    ID As Long
    strID As String
End Type

Public Type Type_Param
    Value As String
    strID As String
End Type

Public Type Type_ResourceInSound
    ID As Long
    strID As String
    Unknown As Long
End Type

Type Type_Sound
    ID As Long
    sndName As String
    Flags As String
    
    ResourceCount As Long
    Resource() As Type_ResourceInSound
    
    Edit As Boolean
End Type

Type Type_SoundResource
    ID As Long
    sndName As String
    Flags As String
    
    Edit As Boolean
End Type

Type Type_Troops
    '#  Each troop contains the following fields:
    '#  1) Troop id (string): used for referencing troops in other files. The prefix trp_ is automatically added before each troop-id .
    '#  2) Toop name (string).
    '#  3) Plural troop name (string).
    '#  4) Troop flags (int). See header_troops.py for a list of available flags
    '#  5) Scene (int) (only applicable to heroes) For example: scn_reyvadin_castle|entry(1) puts troop in reyvadin castle's first entry point
    '#  6) Reserved (int). Put constant "reserved" or 0.
    '#  7) Faction (int)
    '#  8) Inventory (list): Must be a list of items
    '#  9) Attributes (int): Example usage:
    '#           str_6|agi_6|int_4|cha_5|level(5)
    '# 10) Weapon proficiencies (int): Example usage:
    '#           wp_one_handed(55)|wp_two_handed(90)|wp_polearm(36)|wp_archery(80)|wp_crossbow(24)|wp_throwing(45)
    '#     The function wp(x) will create random weapon proficiencies close to Value x.
    '#     To make an expert archer with other weapon proficiencies close to 60 you can use something like:
    '#           wp_archery(160) | wp(60)
    '# 11) Skills (int): See header_skills.py to see a list of skills. Example:
    '#           knows_ironflesh_3|knows_power_strike_2|knows_athletics_2|knows_riding_2
    '# 12) Face code (int): You can obtain the face code by pressing ctrl+E in face generator screen
    '# 13) Face code (int)(2) (only applicable to regular troops, can be omitted for heroes):
    '#     The game will create random faces between Face code 1 and face code 2 for generated troops
    ID As Long
    strID As String
    strName As String
    strPtName As String
    
    '*1.1x ��1.011��Ĳ���*
    unknown_warband(1 To 1) As String
    
    Flags As String
    
    Scene As Long
    SceneID As Long
    Scene_strID As String
    Entry As Long
    
    reserved As Long
    
    Faction As Long
    Faction_strID As String     '����
    
    Upgrade1 As String
    Upgrade1_strID As String
    Upgrade2 As String
    Upgrade2_strID As String
    
    lstInventory(1 To 64) As Type_XY_Index
    
    tAttrib As type_Attrib
    WP As Type_WeaponProf
    Skills(1 To 6) As String
    Face(1 To 8) As String
    
    Edit As Boolean
    csvName As String
    csvName_pl As String
    
End Type

Type Type_Stacks
    ID As Long
    strID As String      '����
    Min As Long
    Max As Long
    Flags As Long '0:member; 1:��².
End Type

Public Type Type_Op_Block
    Op As String
    ParaNum As Long
    Para() As Type_Param
End Type

Type Type_Trigger
    tiOn As Double
    ActNum As Long
    tiAct() As Type_Op_Block
End Type

Type Type_PT
    '#  Each party template record contains the following fields:
    '#  1) Party-template id: used for referencing party-templates in other files.
    '#     The prefix pt_ is automatically added before each party-template id.
    '#  2) Party-template name.
    '#  3) Party flags. See header_parties.py for a list of available flags
    '#  4) Menu. ID of the menu to use when this party is met. The Value 0 uses the default party encounter system.
    '#  5) Faction
    '#  6) Personality. See header_parties.py for an explanation of personality flags.
    '#  7) List of stacks. Each stack record is a tuple that contains the following fields:
    '#    7.1) Troop-id.
    '#    7.2) Minimum number of troops in the stack.
    '#    7.3) Maximum number of troops in the stack.
    '#    7.4) Member flags(optional). Use pmf_is_prisoner to note that this member is a prisoner.
    '#     Note: There can be at most 6 stacks.
    ID As Long
    ptID As String
    ptName As String
    Flags As String
    Menu As String
    Faction As Long
    Faction_strID As String
    
    Personality As Long
    
    Stacks(1 To 6) As Type_Stacks
    
    Edit As Boolean
    csvName As String
    
End Type

Type Type_Party
'####################################################################################################################
'#  Each party record contains the following fields:
'#  1) Party id: used for referencing parties in other files.
'#     The prefix p_ is automatically added before each party id.
'#  2) Party name.
'#  3) Party flags. See header_parties.py for a list of available flags
'#  4) Menu. ID of the menu to use when this party is met. The value 0 uses the default party encounter system.
'#  5) Party-template. ID of the party template this party belongs to. Use pt_none as the default value.
'#  6) Faction.
'#  7) Personality. See header_parties.py for an explanation of personality flags.
'#  8) Ai-behavior
'#  9) Ai-target party
'# 10) Initial coordinates.
'# 11) List of stacks. Each stack record is a triple that contains the following fields:
'#   11.1) Troop-id.
'#   11.2) Number of troops in this stack.
'#   11.3) Member flags. Use pmf_is_prisoner to note that this member is a prisoner.
'# 12) Party direction in degrees [optional]
'####################################################################################################################
    UnknownTitle As Long
    ID As Long
    id2 As Long
    strID As String
    strName As String
    
    Flags As String
    MapIcon_strID As String
    
    Menu As String
    
    Template As Long
    Template_strID As String
    
    Faction As Long
    Faction_strID As String
    
    Personality(1 To 2) As Long
    AI_Behavior As String
    AI_Target As String
    AI_Target_strID As String
    
    reserved As Long
    InitPos(1 To 3) As Type_strXY
    
    UnknownStr As String
    StacksCount As Long
    Stacks() As Type_Stacks
    Degree As String        '��λ:��
    
    Edit As Boolean
    csvName As String
    
End Type

Type Type_Item
    ID As Long '��Ʒ���
    dbName As String '��Ʒ�����ݿ��е�����
    disname As String '��Ʒ����Ϸ����ʾ������
    texname As String '��ͼ������
    nmdl As Long 'ģ�͵�������ͨ����һ��ģ���Ǳ��壬�����ǽ���
    mdlname() As String 'ģ�͵�����
    mdl_b() As String 'ģ�Ͳ���
    'container(3) As String     '���ʵ�����
    'container_binary(3) As String '����λ�õĴ���
    itmType As String '��Ʒ����
    Action As String '��������
    price As Long '�۸�
    'prefix As Long  'ǰ׺���� v0.951
    Prefix As String  'ǰ׺���� v0.952
    'weight As Double '����  'v0.951
    weight As String '����  'v0.952
    abundance As Long '��ԣ��
    head_armor As Long
    body_armor As Long
    leg_armor As Long
    difficulty As Long
    hit_points As Long
    speed_rating As Long
    missile_speed As Long
    weapon_length As Long
    max_ammo As Long
    thrust_damage As Long
    swing_damage As Long
    
    FactionCount As Long
    Faction() As Type_Chest      '������Ӫ
    
    TriggerCount As Long
    Trigger() As Type_Trigger  '���ڻ�ǹ�Ĵ���
    
    Edit As Boolean
    csvName As String
    csvName_pl As String
End Type

Type Type_Faction
    ID As Long
    strID As String
    strName As String
    Flags As String
    lColor As String
    'RelationShip() As Double
    RelationShip() As Type_RelationShip
    
    reserved As Long
    
    csvName As String
    Edit As Boolean
End Type

Type Type_ImodBits
    ID As String
    csvName As String
End Type

Public Type Type_MapIcon
    ID As Long
    strID As String
    Flags As Long
    MeshName As String
    mScale As String
    Sound As Long
    Sound_sndName As String
    
    Offset(0 To 2) As String
    
    TriggerCount As Long
    Triggers() As Type_Trigger  '������
    
    Edit As Boolean
End Type

Public Type Type_TerrainInfo
    Code As String
    Length As Long
End Type

Type Type_Scene
     ID As Long
     strID As String
     strName As String
     Flags As String
     MeshName As String
     BodyName As String
     p(0 To 1) As tPoint
     WaterLevel As Double
     TerrainCode As String
     AccessCount As Long
     Accesses() As String
     
     ChestCount As Long
     Chests() As Type_Chest
     Outer_Terrain_Type As String
     
     Edit As Boolean
End Type

Type Double_XY
     X As Double
     Y As Double
End Type

Type Type_Particle_System
'  1) Particle system id (string)
'  2) Particle system flags (int). See header_particle_systems.py for a list of available flags
'  3) mesh-name.
''''
'  4) Num particles per second:    Number of particles emitted per second.
'  5) Particle Life:    Each particle lives this long (in seconds).
'  6) Damping:          How much particle's speed is lost due to friction.
'  7) Gravity strength: Effect of gravity. (Negative values make the particles float upwards.)
'  8) Turbulance size:  Size of random turbulance (in meters)
'  9) Turbulance strength: How much a particle is affected by turbulance.
''''
' 10,11) Alpha keys :    Each attribute is controlled by two keys and
' 12,13) Red keys   :    each key has two fields: (time, magnitude)
' 14,15) Green keys :    For example scale key (0.3,0.6) means
' 16,17) Blue keys  :    scale of each particle will be 0.6 at the
' 18,19) Scale keys :    time 0.3 (where time=0 means creation and time=1 means end of the particle)
'
' The magnitudes are interpolated in between the two keys and remain constant beyond the keys.
' Except the alpha always starts from 0 at time 0.
''''
' 20) Emit Box Size :   The dimension of the box particles are emitted from.
' 21) Emit velocity :   Particles are initially shot with this velocity.
' 22) Emit dir randomness
' 23) Particle rotation speed: Particles start to rotate with this (angular) speed (degrees per second).
' 24) Particle rotation damping: How quickly particles stop their rotation
    ID As Long
    strID As String
    Flags As String
    Mesh_Name As String
    Particles_Num As Long
    Life As Double
    Damping As Double
    Gravity As Double
    Turbulance_SZ As Double
    Turbulance_Str As Double
    
    Alphak(1) As Double_XY
    Redk(1) As Double_XY
    Greenk(1) As Double_XY
    Bluek(1) As Double_XY
    Scalek(1) As Double_XY
    
    EBSZ(2) As Double
    EV(2) As Double
    EDR As Double
    PRS As Double
    PRD As Double
    
    Edit As Boolean
End Type

Type Type_Tableau_Material
'#######################################################################################################################
'#  1) Tableau id (string)
'#  2) Tableau flags (int)
'#  3) Tableau sample material name (string).
'#  4) Tableau width (int).
'#  5) Tableau height (int).
'#  6) Tableau mesh min x (int): divided by 1000 and used when a mesh is auto-generated using the tableau material
'#  7) Tableau mesh min y (int): divided by 1000 and used when a mesh is auto-generated using the tableau material
'#  8) Tableau mesh max x (int): divided by 1000 and used when a mesh is auto-generated using the tableau material
'#  9) Tableau mesh max y (int): divided by 1000 and used when a mesh is auto-generated using the tableau material
'#  10) Operations block (list): A list of operations
'#     The operations block is executed when the tableau is activated.
'#######################################################################################################################
   ID As Long
   strID As String
   Flags As String
   Sample As String
   Width As Long
   Height As Long
   Min As Type_XY
   Max As Type_XY
   OpCount As Long
   OpBlock() As Type_Op_Block
   Edit As Boolean
End Type

Type Type_Mesh
'####################################################################################################################
'#  Each mesh record contains the following fields:
'#  1) Mesh id: used for referencing meshes in other files. The prefix mesh_ is automatically added before each mesh id.
'#  2) Mesh flags. See header_meshes.py for a list of available flags
'#  3) Mesh resource name: Resource name of the mesh
'#  4) Mesh translation on x axis: Will be done automatically when the mesh is loaded
'#  5) Mesh translation on y axis: Will be done automatically when the mesh is loaded
'#  6) Mesh translation on z axis: Will be done automatically when the mesh is loaded
'#  7) Mesh rotation angle over x axis: Will be done automatically when the mesh is loaded
'#  8) Mesh rotation angle over y axis: Will be done automatically when the mesh is loaded
'#  9) Mesh rotation angle over z axis: Will be done automatically when the mesh is loaded
'#  10) Mesh x scale: Will be done automatically when the mesh is loaded
'#  11) Mesh y scale: Will be done automatically when the mesh is loaded
'#  12) Mesh z scale: Will be done automatically when the mesh is loaded
'####################################################################################################################
    ID As Long
    strID As String
    Flags As String
    Resource_Name As String
    Translation As Type_dblXYZ
    Rotation_Angle As Type_dblXYZ
    Scale As Type_dblXYZ
    Edit As Boolean
End Type

Type Type_Time_Trigger
'####################################################################################################################
'#  Each trigger contains the following fields:
'# 1) Check interval: How frequently this trigger will be checked
'# 2) Delay interval: Time to wait before applying the consequences of the trigger
'#    After its conditions have been evaluated as true.
'# 3) Re-arm interval. How much time must pass after applying the consequences of the trigger for the trigger to become active again.
'#    You can put the constant ti_once here to make sure that the trigger never becomes active again after it fires once.
'# 4) Conditions block (list). This must be a valid operation block. See header_operations.py for reference.
'#    Every time the trigger is checked, the conditions block will be executed.
'#    If the conditions block returns true, the consequences block will be executed.
'#    If the conditions block is empty, it is assumed that it always evaluates to true.
'# 5) Consequences block (list). This must be a valid operation block. See header_operations.py for reference.
'####################################################################################################################
    ID As Long
    Check_Interval As Double
    Delay_Interval As Double
    Rearm_Interval As Double
    Condition() As Type_Op_Block
    ConditionsCount As Long
    Consequence() As Type_Op_Block
    ConsequencesCount As Long
    Edit As Boolean
End Type

Type Type_Global_Variable
    ID As Long
    VarName As String
    'Uses as Integer
End Type

Type Type_String
    ID As Long
    Name As String
    Str As String
    CSV As String
    Edit As Boolean
End Type

Public ShortTags(1 To 26) As Type_XY

'A_E
Public Trps() As Type_Troops
Public PTs() As Type_PT
Public Parties() As Type_Party
Public itm() As Type_Item
Public Factions() As Type_Faction
Public Scenes() As Type_Scene
Public itmID() As Long
Public IMod() As Type_ImodBits
Public PSys() As Type_Particle_System
Public MapIcons() As Type_MapIcon
Public Sounds() As Type_Sound
Public SoundRess() As Type_SoundResource
Public TabMat() As Type_Tableau_Material
Public Mesh() As Type_Mesh
Public TimeTrg() As Type_Time_Trigger
Public gVars() As Type_Global_Variable
Public qStrs() As Type_String
Public Strs() As Type_String

Public ItmVersionInform(2) As String
Public PTVersionInform(2) As String
Public PartyVersionInform(2) As String
Public TrpsVersionInform(2) As String
Public FactionVersionInform(2) As String
Public SceneVersionInform(2) As String
Public PSysVersionInform(2) As String
Public MapIconVersionInform(2) As String
Public SoundVersionInform(2) As String
Public TabMatVersionInform(2) As String
Public MeshVersionInform(2) As String
Public TimeTrgVersionInform(2) As String
Public StringVersionInform(0) As String

'A_E
Public N_Item As Long 'װ������
Public N_Troop As Long 'troop����
Public N_PT As Long 'pt����
Public N_Party As Long 'party����
Public N_Party2 As Long 'party����
Public N_Faction As Long '��Ӫ����
Public N_Scene As Long '��������
Public N_IMod As Long '��Ʒǰ׺����
Public N_PSys As Long '����ϵͳ����
Public N_MapIcon As Long '���ͼͼ������
Public N_Sound As Long '��������
Public N_SoundRes As Long '������Դ����
Public N_TabMat As Long '�ɱ��ز�����
Public N_Mesh As Long '����ģ������
Public N_TimeTrg As Long   '����������
Public N_gVar As Long   'ȫ�ֱ�������
Public N_qStr As Long   '�����ַ�������
Public N_Str As Long    '�ַ�������

'A_E
Public Pointer As Long '���ļ�ʱ��ָ��
Public LinePointer As Long '���ı���ʱ��ָ��
Public CurrentItmID As Long '��ǰװ����
Public CurrentItm As Type_Item '��ǰitem����
Public CurrentTrpID As Long '��ǰtroop��
Public CurrentTrp As Type_Troops '��ǰtroop����
Public CurrentFactionID As Long '��ǰfaction��
Public CurrentFaction As Type_Faction  '��ǰfaction����
Public CurPartyTemplateID As Long '��ǰPartyTemplate��.
Public CurPartyTemplate As Type_PT   '��ǰPartyTemplate����
Public CurrentPartyID As Long '��ǰParty��.
Public CurrentParty As Type_Party   '��ǰParty����
Public CurrentSceneID As Long '��ǰScene��.
Public CurrentScene As Type_Scene   '��ǰScene����
Public CurrentMapIconID As Long '��ǰMapIcon��.
Public CurrentMapIcon As Type_MapIcon  '��ǰMapIcon����
Public CurrentSoundID As Long '��ǰSound��.
Public CurrentSound As Type_Sound  '��ǰSound����
Public CurrentSoundResID As Long '��ǰSoundResource��.
Public CurrentSoundRes As Type_SoundResource   '��ǰSoundResource����
Public CurrentPSysID As Long    '��ǰParticles System��
Public CurrentPSys As Type_Particle_System   '��ǰParticles System����
Public CurrentTabMatID As Long    '��ǰ�ɱ��زĺ�
Public CurrentTabMat As Type_Tableau_Material    '��ǰ�ɱ��زĿ���
Public CurrentMeshID As Long    '��ǰ����ģ�ͺ�
Public CurrentMesh As Type_Mesh    '��ǰ����ģ�Ϳ���
Public CurrentTimeTrgID As Long    '��ǰ��������
Public CurrentTimeTrg As Type_Time_Trigger     '��ǰ����������
'Public CurrentqStrID As Long    '��ǰ�����ַ�����
'Public CurrentqStr As Type_Quick_String     '��ǰ�����ַ�������
Public CurrentStrID As Long    '��ǰ�ַ�����
Public CurrentStr As Type_String     '��ǰ�ַ�������

Public TempItemTrigger As Type_Trigger    '��ʱ��Ʒtrigger,��������
Public TempAct As Type_Op_Block    '��ʱact,��������

Public itmEditCount As Long
Public tpsEditCount As Long
Public tpsItmTabsFrameIndex As Integer '��ǰ��ʾ��

Public bigNumArray() As Byte
Public MaxPointer As Long '���ָ��λ��.

Public lngHandle As Long '�ļ����
Public txtLine As String '�ı�����

Public i64b_num1 As Integer64b
Public i64b_num2 As Integer64b

Public CountDream As Integer
Public DreamTeam(0 To 99) As Integer

Public gBackupFileFlag As Boolean

Public gSkillName(0 To 41) As String
Public gTroopsFlagChk(0 To 63) As String * 12

Public Type saveFlag
    changeAllPartyNumber As Boolean
    ptMax As Double
    ptMin As Double
End Type
Public gSaveFlag As saveFlag

'====== Refactoring ======
'Public gIniFileName As String
'Public gModIniFileName As String
'Public gStrModPath As String
'Public gModName As String
'Public gCSVPath As String 'CSV �ļ���·��

'�洢��ֵΪԭMOD�������ݵ����һ�������ID   'A_E
Public Const EditInfo_TroopsCount = 0
Public Const EditInfo_ItemsCount = 1
Public Const EditInfo_ScenesCount = 2
Public Const EditInfo_FactionsCount = 3
Public Const EditInfo_PartyTemplatesCount = 4
Public Const EditInfo_PartiesCount = 5
Public Const EditInfo_MapIconsCount = 6
Public Const EditInfo_SoundsCount = 7
Public Const EditInfo_SoundRessCount = 8
Public Const EditInfo_PSysCount = 9
Public Const EditInfo_TabMatCount = 10
Public Const EditInfo_MeshCount = 11
Public Const EditInfo_TimeTrgCount = 12
Public Const EditInfo_StringsCount = 13

'��ѡ������
Public IsLoadString As Boolean

Type Type_Global_Variable_Symbol
    iniSetting As String
    iniFileName As String
    ModIniFileName As String
    ModPath As String
    ModName As String
    ModBackUp As String
    CVSPath As String
    LastError As String
    Version As String
    Language As String
    Language_Edit As String
    Op_Set As String
    MBHome As String
    MBsaves As String
    MBsets As String
    
    FirstTimeEdit As Boolean
    EditInfo(13) As Long    'A_E
    InfoFinished As Boolean
    InitFinished As Boolean
End Type
Public MnBInfo As Type_Global_Variable_Symbol

Type kiss_type_run_error_define
    MissCSV As Boolean
    MissMod As Boolean
    MissINI As Boolean
End Type
Public RunERR As kiss_type_run_error_define

Public DisplayMode As Long
Public TriggerBoard As Type_Trigger
Public TriggerCopied As Boolean
Public isShowTip As Boolean

'*************************************************************************
'**�� �� ����clearCSVName
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-06-15 17:44:41
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.955.13
'*************************************************************************
Public Sub clearCSVName()

    Dim n As Integer
    For n = 0 To N_Troop - 1
        'trps(n).csvName = trps(n).strName
        Trps(n).csvName = ""
    Next

    For n = 0 To N_Item - 1
        'Itm(n).csvName = Itm(n).disname
        itm(n).csvName = ""
    Next

    For n = 0 To N_PT - 1
        'PTs(n).csvName = PTs(n).ptName
        PTs(n).csvName = ""
    Next
End Sub

'*************************************************************************
'**�� �� ����LoadPartyCSVFile
'**��    �룺FileName(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-08 16:40:15
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Function LoadPartyCSVFile(FileName As String) As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    LoadPartyCSVFile = False
    
    If LCase$(MnBInfo.Language) = "en" Then
         LoadPartyCSVFile = True
         Exit Function
    End If
    
    If RunERR.MissCSV Then
        Exit Function
    End If

    Dim tmpFileName As String
    tmpFileName = FileName

    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�:" & tmpFileName
        Exit Function
    End If

    Dim TemP As String
    Dim arrTmp() As String
    Dim Index As Long
    Dim n As Long
    Index = 0

    Dim arrFileBuff() As String
    Dim i As Long

    TemP = UEFLoadTextFile(tmpFileName, UEF_UTF8)
    If TemP = vbNullString Then
        MsgBox "vbNullString"
        Exit Function
    End If
    arrFileBuff = Split(TemP, vbCrLf)

    For i = 0 To UBound(arrFileBuff)
        TemP = arrFileBuff(i)
        If Len(Trim$(TemP)) < 1 Then
            TemP = ""
        Else
            arrTmp = Split(TemP, "|")

            For n = 0 To N_Party - 1
                If LCase$(Parties(n).strID) = LCase$(arrTmp(0)) Then
                    Parties(n).csvName = arrTmp(1)
                    'Exit For
                End If
            Next
        End If
    Next
    LoadPartyCSVFile = True

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadPartyCSVFile:[" & FileName & "]", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����LoadPartyTemplateCSVFile
'**��    �룺FileName(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:22:44
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Function LoadPartyTemplateCSVFile(FileName As String) As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    LoadPartyTemplateCSVFile = False
    
    If LCase$(MnBInfo.Language) = "en" Then
         LoadPartyTemplateCSVFile = True
         Exit Function
    End If
    
    If RunERR.MissCSV Then
        Exit Function
    End If

    Dim tmpFileName As String
    tmpFileName = FileName

    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�:" & tmpFileName
        Exit Function
    End If

    Dim TemP As String
    Dim arrTmp() As String
    Dim Index As Long
    Dim n As Long
    Index = 0

    Dim arrFileBuff() As String
    Dim i As Long

    TemP = UEFLoadTextFile(tmpFileName, UEF_UTF8)
    If TemP = vbNullString Then
        MsgBox "vbNullString"
        Exit Function
    End If
    arrFileBuff = Split(TemP, vbCrLf)

    For i = 0 To UBound(arrFileBuff)
        TemP = arrFileBuff(i)
        If Len(Trim$(TemP)) < 1 Then
            TemP = ""
        Else
            arrTmp = Split(TemP, "|")

            For n = 0 To N_PT - 1
                If LCase$(PTs(n).ptID) = LCase$(arrTmp(0)) Then
                    PTs(n).csvName = arrTmp(1)
                    'Exit For
                End If
            Next
        End If
    Next

    'pt_kingdom_1_reinforcements
    'For n = 0 To N_PT - 1
    '    If InStr(PTs(n).ptName, "kingdom") > 0 And InStr(PTs(n).ptName, "reinforcement") > 0 Then
    '        If InStr(PTs(n).csvName, "kingdom") > 0 Then
    '            PTs(n).csvName = "ĳ�����Ĳ���"
    '        End If
   '     End If
   ' Next n

    LoadPartyTemplateCSVFile = True

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadPartyTemplateCSVFile:[" & FileName & "]", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����LoadItemCSVFile
'**��    �룺FileName(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:07
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-11-25 23:55:10
'**��    ����V1.1321
'*************************************************************************
Function LoadItemCSVFile(FileName As String) As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    LoadItemCSVFile = False
    
    If LCase$(MnBInfo.Language) = "en" Then
         LoadItemCSVFile = True
         Exit Function
    End If
    
    If RunERR.MissCSV Then
        Exit Function
    End If
    
    Dim tmpFileName As String
    tmpFileName = FileName
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Function
    End If
    
    Dim TemP As String
    Dim arrTmp() As String
    Dim Index As Long
    Dim n As Long
    Index = 0

    Dim arrFileBuff() As String
    Dim i As Long

    TemP = UEFLoadTextFile(tmpFileName, UEF_UTF8)
    If TemP = vbNullString Then
        MsgBox "vbNullString"
        Exit Function
    End If
    arrFileBuff = Split(TemP, vbCrLf)

    For i = 0 To UBound(arrFileBuff)
        TemP = arrFileBuff(i)
        If Len(Trim$(TemP)) < 1 Then
            TemP = ""
        Else
            arrTmp = Split(TemP, "|")

            For n = 0 To N_Item - 1
                If LCase$(itm(n).dbName) = LCase$(arrTmp(0)) Then
                    itm(n).csvName = arrTmp(1)
                    'Exit For
                End If
            
            
            If Right(arrTmp(0), 3) = "_pl" Then
                     If LCase$(itm(n).dbName) = LCase$(Left(arrTmp(0), Len(arrTmp(0)) - 3)) Then
                        itm(n).csvName_pl = arrTmp(1)
                        'Exit For
                     End If
            End If
            
            Next n
        End If
    Next

    LoadItemCSVFile = True

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadItemCSVFile:[" & FileName & "]", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����LoadTroopCSVFile
'**��    �룺FileName(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:10
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-11-22 23:36:58
'**��    ����V1.1321
'*************************************************************************
Function LoadTroopCSVFile(FileName As String) As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    If LCase$(MnBInfo.Language) = "en" Then
         LoadTroopCSVFile = True
         Exit Function
    End If
    
    If RunERR.MissCSV Then
        Exit Function
    End If

    LoadTroopCSVFile = False
    Dim tmpFileName As String
    tmpFileName = FileName
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Function
    End If

    Dim TemP As String
    Dim arrTmp() As String
    Dim Index As Long
    Dim n As Long
    Index = 0

    Dim arrFileBuff() As String
    Dim i As Long

    TemP = UEFLoadTextFile(tmpFileName, UEF_UTF8)
    If TemP = vbNullString Then
        MsgBox "vbNullString"
        Exit Function
    End If
    arrFileBuff = Split(TemP, vbCrLf)

    For i = 0 To UBound(arrFileBuff)
        TemP = arrFileBuff(i)
        If Len(Trim$(TemP)) < 1 Then
            TemP = ""
        Else
            arrTmp = Split(TemP, "|")

            For n = 0 To N_Troop - 1
                'If Len(trps(n).csvName) = 0 Then
                If LCase$(Trps(n).strID) = LCase$(arrTmp(0)) Then
                    Trps(n).csvName = arrTmp(1)
                    'Exit For
                End If
                
                If Right(arrTmp(0), 3) = "_pl" Then
                     If LCase$(Trps(n).strID) = LCase$(Left(arrTmp(0), Len(arrTmp(0)) - 3)) Then
                        Trps(n).csvName_pl = arrTmp(1)
                        'Exit For
                     End If
                End If
                'End If
            Next
        End If
    Next

    LoadTroopCSVFile = True

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadTroopCSVFile:[" & FileName & "]", Err.Number, Err.Description)
End Function
'*************************************************************************
'**�� �� ����LoadIModCSVFile
'**��    �룺FileName(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:07
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-11-25 23:55:10
'**��    ����V1.1321
'*************************************************************************
Function LoadIModCSVFile(FileName As String) As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim i As Long

    LoadIModCSVFile = False
    
    If LCase$(MnBInfo.Language) = "en" Then
         LoadIModCSVFile = True
         Exit Function
    End If
    
    If RunERR.MissCSV Then
        Exit Function
    End If
    
    Dim tmpFileName As String
    tmpFileName = FileName
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Function
    End If
    
    Dim TemP As String
    Dim arrTmp() As String
    Dim Index As Long
    Dim n As Long
    Index = 0

    Dim arrFileBuff() As String

    TemP = UEFLoadTextFile(tmpFileName, UEF_UTF8)
    If TemP = vbNullString Then
        MsgBox "vbNullString"
        Exit Function
    End If
    
    arrFileBuff = Split(TemP, vbCrLf)
    
     For i = 0 To N_IMod - 1
        For n = 0 To UBound(arrFileBuff)
         TemP = arrFileBuff(n)
         If Len(Trim$(TemP)) < 1 Then
         TemP = ""
         Else
         arrTmp = Split(TemP, "|")
         End If
            If LCase(Trim(IMod(i).ID)) = LCase(Trim(arrTmp(0))) Then
                arrTmp(1) = Replace(arrTmp(1), "%s", "")
                arrTmp(1) = Trim(arrTmp(1))
                IMod(i).csvName = arrTmp(1)
                Exit For
            End If
        Next n
        
    Next i

    LoadIModCSVFile = True

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadIModCSVFile:[" & FileName & "]", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����LoadFactionCSVFile
'**��    �룺FileName(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:10
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-11-22 23:36:58
'**��    ����V1.1321
'*************************************************************************
Function LoadFactionCSVFile(FileName As String) As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    If RunERR.MissCSV Then
        Exit Function
    End If

    LoadFactionCSVFile = False
    Dim tmpFileName As String
    tmpFileName = FileName
    
    If LCase$(MnBInfo.Language) = "en" Then
         LoadFactionCSVFile = True
         Exit Function
    End If
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Function
    End If

    Dim TemP As String
    Dim arrTmp() As String
    Dim Index As Long
    Dim n As Long
    Index = 0

    Dim arrFileBuff() As String
    Dim i As Long

    TemP = UEFLoadTextFile(tmpFileName, UEF_UTF8)
    If TemP = vbNullString Then
        MsgBox "vbNullString"
        Exit Function
    End If
    arrFileBuff = Split(TemP, vbCrLf)

    For i = 0 To UBound(arrFileBuff)
        TemP = arrFileBuff(i)
        If Len(Trim$(TemP)) < 1 Then
            TemP = ""
        Else
            arrTmp = Split(TemP, "|")

            For n = 0 To N_Faction - 1
                'If Len(Factionss(n).csvName) = 0 Then
                If LCase$(Factions(n).strID) = LCase$(arrTmp(0)) Then
                    Factions(n).csvName = arrTmp(1)
                    'Exit For
                End If
                
                'End If
            Next
        End If
    Next

    LoadFactionCSVFile = True

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadFactionCSVFile:[" & FileName & "]", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����LoadQuickStringCSVFile
'**��    �룺FileName(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2012-03-02 22:07:10
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Function LoadQuickStringCSVFile(FileName As String) As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    LoadQuickStringCSVFile = False
    
    If LCase$(MnBInfo.Language) = "en" Then
         LoadQuickStringCSVFile = True
         Exit Function
    End If
    
    If RunERR.MissCSV Then
        Exit Function
    End If
    
    Dim tmpFileName As String
    tmpFileName = FileName
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Function
    End If
    
    Dim TemP As String
    Dim arrTmp() As String
    Dim Index As Long
    Dim n As Long
    Index = 0

    Dim arrFileBuff() As String
    Dim i As Long

    TemP = UEFLoadTextFile(tmpFileName, UEF_UTF8)
    If TemP = vbNullString Then
        MsgBox "vbNullString"
        Exit Function
    End If
    arrFileBuff = Split(TemP, vbCrLf)
    For i = 0 To UBound(arrFileBuff)
        TemP = arrFileBuff(i)
        If Len(Trim$(TemP)) < 1 Then
            TemP = ""
        Else
            arrTmp = Split(TemP, "|")

            For n = 0 To N_qStr - 1
                If LCase$(qStrs(n).Name) = LCase$(arrTmp(0)) Then
                    qStrs(n).CSV = arrTmp(1)
                End If
            Next
        End If
    Next
    LoadQuickStringCSVFile = True

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadQuickStringCSVFile:[" & FileName & "]", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����LoadStringCSVFile
'**��    �룺FileName(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2012-03-02 22:37:14
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Function LoadStringCSVFile(FileName As String) As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    LoadStringCSVFile = False
    
    If LCase$(MnBInfo.Language) = "en" Then
         LoadStringCSVFile = True
         Exit Function
    End If
    
    If RunERR.MissCSV Then
        Exit Function
    End If
    
    Dim tmpFileName As String
    tmpFileName = FileName
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Function
    End If
    
    Dim TemP As String
    Dim arrTmp() As String
    Dim Index As Long
    Dim n As Long, H As Long
    Dim F As Boolean
    
    Index = 0

    Dim arrFileBuff() As String
    Dim i As Long

    TemP = UEFLoadTextFile(tmpFileName, UEF_UTF8)
    If TemP = vbNullString Then
        MsgBox "vbNullString"
        Exit Function
    End If
    arrFileBuff = Split(TemP, vbCrLf)
    H = 0
    For i = 0 To UBound(arrFileBuff)
        TemP = arrFileBuff(i)
        If Len(Trim$(TemP)) < 1 Then
            TemP = ""
        Else
            arrTmp = Split(TemP, "|")
            F = False
            For n = H To N_Str - 1
                If LCase$(Strs(n).Name) = LCase$(arrTmp(0)) Then
                    Strs(n).CSV = arrTmp(1)
                    H = n + 1
                    F = True
                    Exit For
                End If
            Next n
            
            If Not F Then
               For n = 0 To H - 1
                   If LCase$(Strs(n).Name) = LCase$(arrTmp(0)) Then
                      Strs(n).CSV = arrTmp(1)
                      Exit For
                   End If
               Next n
            End If
            
        End If
    Next
    LoadStringCSVFile = True

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadStringCSVFile:[" & FileName & "]", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����GetWord
'**��    �룺��
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:14
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Function GetWord() As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    GetWord = ""
    Dim tmp As Byte
before:
    If Pointer > MaxPointer Then
        Call logErr("ModMain", "GetWord", "OVER_MAX_POINTER", "��������ļ�ָ��!")
        Exit Function
    End If
    Get lngHandle, Pointer, tmp
    Pointer = Pointer + 1

    If tmp = 10 Or tmp = 13 Or tmp = 32 Then
        If GetWord <> "" Then Exit Function
    Else
        GetWord = GetWord & Chr(tmp)
    End If
    GoTo before
    
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "GetWord", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����GetWordL
'**��    �룺(Boolean)bEnd
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2008-05-18 08:23:14
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Function GetWordL(Optional bEnd As Boolean) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    GetWordL = ""
    Dim tmp As Byte
    bEnd = False
before:
    If LinePointer > Len(txtLine) Then
        bEnd = True
        Exit Function
    End If
    tmp = Asc(Mid(txtLine, LinePointer, 1))

    If tmp = 10 Or tmp = 13 Or tmp = 32 Then
        If GetWordL <> "" Then Exit Function
    Else
        GetWordL = GetWordL & Chr(tmp)
    End If
    
    LinePointer = LinePointer + 1
    
    GoTo before
    
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "GetWordL", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����PutWord
'**��    �룺TheWord(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:17
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Sub PutWord(TheWord As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim tmp As Byte
    Dim n As Integer
    
    
    For n = 1 To Len(TheWord)
        tmp = Asc(Mid(TheWord, n, 1))
        If tmp = 32 And n = 1 Then
            n = 2
            tmp = Asc(Mid(TheWord, n, 1))
        End If
        Put lngHandle, Pointer, tmp
        Pointer = Pointer + 1
    Next n
    

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "PutWord", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**�� �� ����PutSpc
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:20
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Sub PutSpc()
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim Mspc As Byte

    Mspc = 32
    Put lngHandle, Pointer, Mspc
    Pointer = Pointer + 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "PutSpc", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**�� �� ����PutReturn
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:25
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Sub PutReturn()
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim tmp As Byte
   
    tmp = 13
    Put lngHandle, Pointer, tmp
    Pointer = Pointer + 1
    tmp = 10
    Put #lngHandle, Pointer, tmp
    Pointer = Pointer + 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "PutReturn", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**�� �� ����GetLine
'**��    �룺��
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:28
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Function GetLine() As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    GetLine = ""
    Dim tmp As Byte
    
   
before:
    Get lngHandle, Pointer, tmp
    Pointer = Pointer + 1
    If Pointer > MaxPointer Then
        Exit Function
    End If
    If tmp = 13 Or tmp = 10 Then
        If GetLine <> "" Then Exit Function
    Else
        GetLine = GetLine & Chr(tmp)
    End If
    GoTo before

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "GetLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����GetRealLine
'**��    �룺��
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:31
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Function GetRealLine() As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    GetRealLine = ""
    Dim tmp As Byte

before:
    Get lngHandle, Pointer, tmp
    Pointer = Pointer + 1
    If Pointer > MaxPointer Then
        Exit Function
    End If
    'If tmp = 13 Or tmp = 10 Then
    If tmp = 13 Then
        Get lngHandle, Pointer, tmp
        Pointer = Pointer + 1
        Exit Function
    Else
        GetRealLine = GetRealLine & Chr(tmp)
    End If
    GoTo before

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "GetRealLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����MinusIT
'**��    �룺thenum(String) -
'**��    ����(Long) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:37
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Function MinusIT(thenum As String) As Long
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim n1, n2 As String
    Dim Num1, Num2 As Long
  
    n1 = "2000000000"
    n2 = Right(thenum, 9)
    Num1 = Val(n1)
    Num2 = Val(n2)
    Num1 = Num1 - 2147483647 - 1
    MinusIT = Num1 + Num2

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "MinusIT", Err.Number, Err.Description)
End Function



'*************************************************************************
'**�� �� ����LoadTroopFile
'**��    �룺filePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:40
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-11-22 23:39:02
'**��    ����V1.1321
'*************************************************************************
Sub LoadTroopFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim str_tmp1 As String
    Dim str_tmp2 As String
    Dim n As Long
    Dim m As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    MaxPointer = FileLen(tmpFileName)
    
    lngHandle = FreeFile()

    Open tmpFileName For Random Access Read As lngHandle Len = 1
    Pointer = 1
    For n = 0 To 2
        TrpsVersionInform(n) = GetWord()
    Next n
    N_Troop = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_TroopsCount) = N_Troop
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_TroopsCount, CStr(N_Troop)
    End If
    
    ReDim Trps(N_Troop - 1)
    
    DoEvents
    For n = 0 To N_Troop - 1

        With Trps(n)
            .ID = n
            .strID = GetWord()
            .strName = GetWord()
            .strPtName = GetWord()
            .csvName = .strName 'default
            .csvName_pl = .strPtName 'default
            .unknown_warband(1) = GetWord()
            
            .Flags = GetWord()
            .Scene = CLng(Val(GetWord()))
            .reserved = Val(GetWord())
            
            .Faction = GetWord()
            
            .Upgrade1 = GetWord()
            .Upgrade2 = GetWord()
            
            For m = 1 To 64
                str_tmp1 = GetWord()
                str_tmp2 = GetWord()
                .lstInventory(m).X = Val(str_tmp1)
                .lstInventory(m).Y = Val(str_tmp2)
            Next m
            
            .tAttrib.strPoint = Val(GetWord())
            .tAttrib.agiPoint = Val(GetWord())
            .tAttrib.intPoint = Val(GetWord())
            .tAttrib.chaPoint = Val(GetWord())
            .tAttrib.level = Val(GetWord())
            
            .WP.one_handed = GetWord()
            .WP.two_handed = GetWord()
            .WP.polearm = GetWord()
            .WP.archery = GetWord()
            .WP.crossbow = GetWord()
            .WP.throwing = GetWord()
            .WP.firearm = GetWord()
            
            .Skills(1) = Val(GetWord())
            .Skills(2) = Val(GetWord())
            .Skills(3) = Val(GetWord())
            .Skills(4) = Val(GetWord())
            .Skills(5) = Val(GetWord())
            .Skills(6) = Val(GetWord())
            
            .Face(1) = GetWord()
            .Face(2) = GetWord()
            .Face(3) = GetWord()
            .Face(4) = GetWord()
            .Face(5) = GetWord()
            .Face(6) = GetWord()
            .Face(7) = GetWord()
            .Face(8) = GetWord()
            
            .Edit = CheckEditable(EditInfo_TroopsCount, n)
            
            AddIndex .ID, .strID
            
        End With
    Next
    Close lngHandle
    Pointer = 1
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadTroopFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����ReadTroopLine
'**��    �룺Text(String),OutputTroop(Type_Troops)
'**��    ����-
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-09 16:24:35
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub ReadTroopLine(Text As String, OutputTroop As Type_Troops)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim str_tmp1 As String
    Dim str_tmp2 As String
    Dim n As Long
    Dim m As Integer
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
        With OutputTroop
            .strID = GetWordL()
            .strName = GetWordL()
            .strPtName = GetWordL()
            .unknown_warband(1) = GetWordL()
            
            .Flags = GetWordL()
            
            .Scene = CLng(Val(GetWordL()))
            PurseSceneInTroop .Scene, .SceneID, .Scene_strID, .Entry
            
            .reserved = Val(GetWordL())
            
            .Faction = GetWordL()
            .Faction_strID = Factions(.Faction).strID
            
            .Upgrade1 = GetWordL()
            .Upgrade1_strID = Trps(.Upgrade1).strID
            .Upgrade2 = GetWordL()
            .Upgrade2_strID = Trps(.Upgrade2).strID
            
            For m = 1 To 64
                str_tmp1 = GetWordL()
                str_tmp2 = GetWordL()
                .lstInventory(m).X = Val(str_tmp1)
                If .lstInventory(m).X > -1 Then
                  .lstInventory(m).strX = itm(.lstInventory(m).X).dbName
                Else
                  .lstInventory(m).strX = ""
                End If
                .lstInventory(m).Y = Val(str_tmp2)
            Next m
            
            .tAttrib.strPoint = Val(GetWordL())
            .tAttrib.agiPoint = Val(GetWordL())
            .tAttrib.intPoint = Val(GetWordL())
            .tAttrib.chaPoint = Val(GetWordL())
            .tAttrib.level = Val(GetWordL())
            
            .WP.one_handed = GetWordL()
            .WP.two_handed = GetWordL()
            .WP.polearm = GetWordL()
            .WP.archery = GetWordL()
            .WP.crossbow = GetWordL()
            .WP.throwing = GetWordL()
            .WP.firearm = GetWordL()
            
            .Skills(1) = Val(GetWordL())
            .Skills(2) = Val(GetWordL())
            .Skills(3) = Val(GetWordL())
            .Skills(4) = Val(GetWordL())
            .Skills(5) = Val(GetWordL())
            .Skills(6) = Val(GetWordL())
            
            .Face(1) = GetWordL()
            .Face(2) = GetWordL()
            .Face(3) = GetWordL()
            .Face(4) = GetWordL()
            .Face(5) = GetWordL()
            .Face(6) = GetWordL()
            .Face(7) = GetWordL()
            .Face(8) = GetWordL()
            
        End With
    
    LinePointer = 1
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ReadTroopLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����LoadFactionFile
'**��    �룺filePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-24 23:03:54
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub LoadFactionFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim n As Long
    Dim m As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    MaxPointer = FileLen(tmpFileName)
    
    lngHandle = FreeFile()

    Open tmpFileName For Random Access Read As lngHandle Len = 1
    Pointer = 1
    For n = 0 To 2
        FactionVersionInform(n) = GetWord()
    Next n
    N_Faction = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_FactionsCount) = N_Faction
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_FactionsCount, CStr(N_Faction)
    End If
    
    ReDim Factions(N_Faction - 1)
    
    DoEvents
    For n = 0 To N_Faction - 1
         With Factions(n)
               .ID = n
               .strID = GetWord()
               .strName = GetWord()
               .csvName = .strName 'Default
               .Flags = Val(GetWord())
               .lColor = GetWord()
               
               ReDim Factions(n).RelationShip(N_Faction - 1)
                
               For m = 0 To N_Faction - 1
                     .RelationShip(m).Value = Val(GetWord())
               Next m
               .reserved = Val(GetWord())
               
               .Edit = CheckEditable(EditInfo_FactionsCount, n)
               
               AddIndex .ID, .strID
         End With
    Next
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadFactionFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����ReadFactionLine
'**��    �룺Text(String),OutputFaction(Type_Faction)
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-09 16:40:22
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub ReadFactionLine(Text As String, OutputFaction As Type_Faction)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim n As Long
    Dim m As Integer
    Dim head As String
    
    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
         With OutputFaction
         
               .strID = GetWordL()
               .strName = GetWordL()
               .Flags = Val(GetWordL())
               .lColor = GetWordL()
               
               ReDim Factions(n).RelationShip(N_Faction - 1)
                
               For m = 0 To N_Faction - 1
                     .RelationShip(m).ID = m
                     .RelationShip(m).strID = Factions(m).strID
                     .RelationShip(m).Value = Val(GetWordL())
               Next m
               .reserved = Val(GetWordL())

         End With

    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ReadFactionLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����LoadSceneFile
'**��    �룺filePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-28 23:07:51
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub LoadSceneFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim n As Long
    Dim m As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    MaxPointer = FileLen(tmpFileName)
    
    lngHandle = FreeFile()

    Open tmpFileName For Random Access Read As lngHandle Len = 1
    Pointer = 1
    For n = 0 To 2
        SceneVersionInform(n) = GetWord()
    Next n
    N_Scene = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_ScenesCount) = N_Scene
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_ScenesCount, CStr(N_Scene)
    End If
    
    ReDim Scenes(N_Scene - 1)
    
    DoEvents
    For n = 0 To N_Scene - 1
         With Scenes(n)
               .ID = n
               .strID = GetWord()
               .strName = GetWord()
               .Flags = GetWord()
               .MeshName = GetWord()
               .BodyName = GetWord()
               
               .p(0).X = Val(GetWord())
               .p(0).Y = Val(GetWord())
               .p(1).X = Val(GetWord())
               .p(1).Y = Val(GetWord())
               
               .WaterLevel = Val(GetWord())
               .TerrainCode = GetWord()
               
               .AccessCount = Val(GetWord())
               
               If .AccessCount > 0 Then
                 ReDim .Accesses(1 To .AccessCount)
               
                 For m = 1 To .AccessCount
                  .Accesses(m) = GetWord()
                 Next m
               Else
                 ReDim .Accesses(0)
               End If
               
               .ChestCount = Val(GetWord())
               
               If .ChestCount > 0 Then
                 ReDim .Chests(1 To .ChestCount)
               
                 For m = 1 To .ChestCount
                  .Chests(m).ID = Val(GetWord())
                 Next m
               Else
                 ReDim .Chests(0)
               End If
               
               .Outer_Terrain_Type = GetWord()
               
               .Edit = CheckEditable(EditInfo_ScenesCount, n)
               
               AddIndex .ID, .strID
         End With
    Next
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadSceneFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����ReadSceneLine
'**��    �룺Text(String),OutputScene(Type_Scene)
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-09 16:59:36
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub ReadSceneLine(Text As String, OutputScene As Type_Scene)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim n As Long
    Dim m As Integer
    Dim head As String
    
    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
         With OutputScene
               .strID = GetWordL()
               .strName = GetWordL()
               .Flags = GetWordL()
               .MeshName = GetWordL()
               .BodyName = GetWordL()
               
               .p(0).X = Val(GetWordL())
               .p(0).Y = Val(GetWordL())
               .p(1).X = Val(GetWordL())
               .p(1).Y = Val(GetWordL())
               
               .WaterLevel = Val(GetWordL())
               .TerrainCode = GetWordL()
               
               .AccessCount = Val(GetWordL())
               
               If .AccessCount > 0 Then
                 ReDim .Accesses(1 To .AccessCount)
               
                 For m = 1 To .AccessCount
                  .Accesses(m) = GetWordL()
                 Next m
               Else
                 ReDim .Accesses(0)
               End If
               
               .ChestCount = Val(GetWordL())
               
               If .ChestCount > 0 Then
                 ReDim .Chests(1 To .ChestCount)
               
                 For m = 1 To .ChestCount
                  .Chests(m).ID = Val(GetWordL())
                  .Chests(m).strID = Trps(.Chests(m).ID).strID
                 Next m
               Else
                 ReDim .Chests(0)
               End If
               
               .Outer_Terrain_Type = GetWordL()

         End With
    
    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ReadSceneLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SaveTroopFile
'**��    �룺filePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:45
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-11-25 16:30:41
'**��    ����V1.1321
'*************************************************************************
Sub SaveTroopFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim tmpStr As String
    Dim n As Long, i As Long
    Dim m As Integer

    lngHandle = FreeFile()

    Open FilePath For Output As #lngHandle

    tmpStr = ""
    For n = 0 To 1
        tmpStr = tmpStr & TrpsVersionInform(n) & " "
    Next n
    tmpStr = tmpStr & TrpsVersionInform(2)
    Print #lngHandle, tmpStr
    Print #lngHandle, Trim$(Str$(N_Troop)) & " "
    For n = 0 To N_Troop - 1

        With Trps(n)

            'Check Value
            If Left$(Trps(n).strID, 4) <> "trp_" Then
                Trps(n).strID = "trp_" & Trps(n).strID
            End If

            Print #lngHandle, .strID & " " & .strName & " " & .strPtName & " " & .unknown_warband(1) & " " & .Flags & _
                  " " & CStr(.Scene) & " " & CStr(.reserved) & " " & CStr(.Faction) & " " & .Upgrade1 & " " & .Upgrade2
            tmpStr = "  "
            For m = 1 To 64
                tmpStr = tmpStr & CStr(.lstInventory(m).X) & " " & CStr(.lstInventory(m).Y) & " "
            Next m
            Print #lngHandle, tmpStr

            Print #lngHandle, "  " & CStr(.tAttrib.strPoint) & " " & CStr(.tAttrib.agiPoint) & _
                  " " & CStr(.tAttrib.intPoint) & " " & CStr(.tAttrib.chaPoint) & " " & CStr(.tAttrib.level)

            Print #lngHandle, " " & .WP.one_handed & " " & .WP.two_handed & " " & .WP.polearm & _
                  " " & .WP.archery & " " & .WP.crossbow & " " & .WP.throwing & " " & .WP.firearm

            Print #lngHandle, "" & CStr(.Skills(1)) & " " & CStr(.Skills(2)) & " " & CStr(.Skills(3)) & _
                  " " & CStr(.Skills(4)) & " " & CStr(.Skills(5)) & " " & CStr(.Skills(6)) & " "

            Print #lngHandle, "  " & .Face(1) & " " & .Face(2) & " " & .Face(3) & " " & .Face(4) & " " & _
                  .Face(5) & " " & .Face(6) & " " & .Face(7) & " " & .Face(8) & " "

            Print #lngHandle, ""
            
        End With
    Next n
    Close lngHandle

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveTroopFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SaveTroopCSVFile
'**��    �룺filePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-30 22:08:41
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SaveTroopCSVFile(FilePath As String)
     On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim strTem As String, i As Long, q As Boolean
    
    If LCase$(MnBInfo.Language) = "en" Then
         'SaveTroopCSVFile = True
         Exit Sub
    End If
    
    For i = 0 To N_Troop - 1
       If Trps(i).csvName <> Trps(i).strName Then
          strTem = strTem & Trps(i).strID & "|" & Trps(i).csvName & vbCrLf
       End If
       
       If Trps(i).csvName_pl <> Trps(i).strPtName Then
          strTem = strTem & Trps(i).strID & "_pl" & "|" & Trps(i).csvName_pl & vbCrLf
       End If
    Next i
    
    q = UEFSaveTextFile(FilePath, strTem, False, UEF_UTF8, UEF_UTF8)
    
    If q = False Then
        Call logErr("ModMain", "SaveTroopCSVFile", "INVALID_HANDLE_VALUE", "�޷����ļ�:[" & FilePath & "]")
    End If
     Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveTroopCSVFile", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**�� �� ����LoadPTFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:23:48
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-11-30 14:52:37
'**��    ����V1.1321
'*************************************************************************
Sub LoadPTFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
        
    Dim tmp1 As String
    Dim n As Long
    Dim m As Integer
    Dim tmpFileName As String
    tmpFileName = FilePath
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�:" & tmpFileName
        Exit Sub
    End If
    
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As #lngHandle Len = 1
    Pointer = 1
    MaxPointer = FileLen(tmpFileName)
    For n = 0 To 2
        PTVersionInform(n) = GetWord()
    Next n
    N_PT = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_PartyTemplatesCount) = N_PT
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_PartyTemplatesCount, CStr(N_PT)
    End If
    
    ReDim PTs(N_PT - 1)
    
    DoEvents
    For n = 0 To N_PT - 1
        
        With PTs(n)
            .ID = n
            .ptID = GetWord()
            .ptName = GetWord()
            .Flags = GetWord()
            .Menu = GetWord()
            .Faction = Val(GetWord())
            .Personality = Val(GetWord())
            
            .csvName = .ptName  'default
            
            For m = 1 To 6
                .Stacks(m).ID = Val(GetWord())
                
                If .Stacks(m).ID < 0 Then
                    'For m1 = m To 6
                    '   .stacks(m1).id = -1
                    '  .stacks(m1).Min = -1
                    '  .stacks(m1).Max = -1
                    '  .stacks(m1).flags = ""
                    'Next
                    'm = m + 10
                Else
                    .Stacks(m).Min = Val(GetWord())
                    .Stacks(m).Max = Val(GetWord())
                    .Stacks(m).Flags = Val(GetWord())
                End If
                
            Next
            
            .Edit = CheckEditable(EditInfo_PartyTemplatesCount, n)
            
            AddIndex .ID, .ptID
        End With
        
    Next n
    
    Close #lngHandle
    Pointer = 1

    
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadPTFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����ReadPTLine
'**��    �룺Text(String),(Type_PT)OutputPT
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-09 17:02:26
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub ReadPTLine(Text As String, OutputPT As Type_PT)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
        
    Dim tmp1 As String
    Dim n As Long
    Dim m As Integer
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    head = GetWordL()
    
        With OutputPT
            .ptID = GetWordL()
            .ptName = GetWordL()
            .Flags = GetWordL()
            .Menu = GetWordL()
            
            .Faction = Val(GetWordL())
            .Faction_strID = Factions(.Faction).strID
            
            .Personality = Val(GetWordL())

            For m = 1 To 6
                .Stacks(m).ID = Val(GetWordL())
                
                If .Stacks(m).ID < 0 Then
                    'For m1 = m To 6
                    '   .stacks(m1).id = -1
                    '  .stacks(m1).Min = -1
                    '  .stacks(m1).Max = -1
                    '  .stacks(m1).flags = ""
                    'Next
                    'm = m + 10
                Else
                    .Stacks(m).strID = Trps(.Stacks(m).ID).strID
                    .Stacks(m).Min = Val(GetWordL())
                    .Stacks(m).Max = Val(GetWordL())
                    .Stacks(m).Flags = Val(GetWordL())
                End If
                
            Next m
            
        End With

    LinePointer = 1

    
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ReadPTLine", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**�� �� ����SavePartyTemplateCSVFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-07 15:44:11
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SavePartyTemplateCSVFile(FilePath As String)
     On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim strTem As String, i As Long, q As Boolean
    
    If LCase$(MnBInfo.Language) = "en" Then
         'SavePartyTemplateCSVFile = True
         Exit Sub
    End If
    
    For i = 0 To N_PT - 1
       If PTs(i).csvName <> PTs(i).ptName Then
          strTem = strTem & PTs(i).ptID & "|" & PTs(i).csvName & vbCrLf
       End If
       
    Next i
    
    q = UEFSaveTextFile(FilePath, strTem, False, UEF_UTF8, UEF_UTF8)
    
    If q = False Then
        Call logErr("ModMain", "SavePartyTemplateCSVFile", "INVALID_HANDLE_VALUE", "�޷����ļ�:[" & FilePath & "]")
    End If
     Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SavePartyTemplateCSVFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SavePartyTemplateFile
'**��    �룺FilePath(String) -PT�ļ���·��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-06-07 16:49:36
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-12-07 15:22:37
'**��    ����V1.1321
'*************************************************************************
Sub SavePartyTemplateFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim i As Double
    Dim n As Long
    Dim m As Integer

    Dim tmptxtLine As String

    lngHandle = FreeFile()
    'fix me
    Open FilePath For Output As #lngHandle
    tmptxtLine = ""
    For n = 0 To 1
        tmptxtLine = tmptxtLine & (PTVersionInform(n)) & " "
    Next n
    tmptxtLine = tmptxtLine & (PTVersionInform(2))
    Print #lngHandle, tmptxtLine

    Print #lngHandle, (CStr(N_PT))

    For n = 0 To N_PT - 1
           tmptxtLine = OutputPTLine(n)
           Print #lngHandle, tmptxtLine
    Next n
    Close #lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SavePartyTemplateFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SavePartyCSVFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-08 16:37:32
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SavePartyCSVFile(FilePath As String)
     On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim strTem As String, i As Long, q As Boolean
    
    If LCase$(MnBInfo.Language) = "en" Then
         'SavePartyTemplateCSVFile = True
         Exit Sub
    End If
    
    For i = 0 To N_Party - 1
       If Parties(i).csvName <> Parties(i).strName Then
          strTem = strTem & Parties(i).strID & "|" & Parties(i).csvName & vbCrLf
       End If
       
    Next i
    
    q = UEFSaveTextFile(FilePath, strTem, False, UEF_UTF8, UEF_UTF8)
    
    If q = False Then
        Call logErr("ModMain", "SavePartyCSVFile", "INVALID_HANDLE_VALUE", "�޷����ļ�:[" & FilePath & "]")
    End If
     Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SavePartyCSVFile", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**�� �� ����SaveSoundFile
'**��    �룺FilePath(String) -Sound�ļ���·��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�
'**��    �ڣ�
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2011-01-03 15:40:38
'**��    ����V1.1321
'*************************************************************************
Sub SaveSoundFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim i As Double
    Dim n As Long
    Dim m As Integer

    Dim tmptxtLine As String

    lngHandle = FreeFile()
    'fix me
    Open FilePath For Output As #lngHandle
    tmptxtLine = ""
    For n = 0 To 1
        tmptxtLine = tmptxtLine & (SoundVersionInform(n)) & " "
    Next n
    tmptxtLine = tmptxtLine & (SoundVersionInform(2))
    Print #lngHandle, tmptxtLine
    
    Print #lngHandle, (CStr(N_SoundRes))

    For n = 0 To N_SoundRes - 1
           tmptxtLine = OutputSoundResLine(n)
           Print #lngHandle, tmptxtLine
    Next n
    
    Print #lngHandle, (CStr(N_Sound))

    For n = 0 To N_Sound - 1
           tmptxtLine = OutputSoundLine(n)
           Print #lngHandle, tmptxtLine
    Next n
    Close #lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveSoundFile", Err.Number, Err.DescriSoundion)
End Sub

'*************************************************************************
'**�� �� ����SavePartyFile
'**��    �룺FilePath(String) -Party�ļ���·��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�
'**��    �ڣ�
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-12-09 13:22:42
'**��    ����V1.1321
'*************************************************************************
Sub SavePartyFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim i As Double
    Dim n As Long
    Dim m As Integer

    Dim tmptxtLine As String

    lngHandle = FreeFile()
    'fix me
    Open FilePath For Output As #lngHandle
    tmptxtLine = ""
    For n = 0 To 1
        tmptxtLine = tmptxtLine & (PartyVersionInform(n)) & " "
    Next n
    tmptxtLine = tmptxtLine & (PartyVersionInform(2))
    Print #lngHandle, tmptxtLine

    Print #lngHandle, (CStr(N_Party)) & " " & (CStr(N_Party2))

    For n = 0 To N_Party - 1
           tmptxtLine = OutputPartyLine(n)
           Print #lngHandle, tmptxtLine
    Next n
    Close #lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SavePartyFile", Err.Number, Err.DescriPartyion)
End Sub

'*************************************************************************
'**�� �� ����SaveSceneFile
'**��    �룺FilePath(String) -Scene�ļ���·��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�
'**��    �ڣ�
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-12-09 13:22:42
'**��    ����V1.1321
'*************************************************************************
Sub SaveSceneFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim i As Double
    Dim n As Long
    Dim m As Integer

    Dim tmptxtLine As String

    lngHandle = FreeFile()
    'fix me
    Open FilePath For Output As #lngHandle
    tmptxtLine = ""
    For n = 0 To 1
        tmptxtLine = tmptxtLine & (SceneVersionInform(n)) & " "
    Next n
    tmptxtLine = tmptxtLine & (SceneVersionInform(2))
    Print #lngHandle, tmptxtLine

    Print #lngHandle, " " & (CStr(N_Scene))

    For n = 0 To N_Scene - 1
           tmptxtLine = OutputSceneLine(n)
           Print #lngHandle, tmptxtLine
    Next n
    Close #lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveSceneFile", Err.Number, Err.DescriSceneion)
End Sub


'*************************************************************************
'**�� �� ����LoadPartyFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-07 23:35:16
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub LoadPartyFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
        
    Dim tmp1 As String
    Dim n As Long
    Dim m As Integer
    Dim tmpFileName As String
    tmpFileName = FilePath
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�:" & tmpFileName
        Exit Sub
    End If
    
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As #lngHandle Len = 1
    Pointer = 1
    MaxPointer = FileLen(tmpFileName)
    For n = 0 To 2
        PartyVersionInform(n) = GetWord()
    Next n
    N_Party = Val(GetWord())
    N_Party2 = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_PartiesCount) = N_Party
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_PartiesCount, CStr(N_Party)
    End If
    
    ReDim Parties(N_Party - 1)
    
    DoEvents
    For n = 0 To N_Party - 1
        
        With Parties(n)
            .UnknownTitle = Val(GetWord())
            .ID = Val(GetWord())
            .id2 = Val(GetWord())
            .strID = GetWord()
            .strName = GetWord()
            .Flags = GetWord()
            .Menu = GetWord()
            .Template = Val(GetWord())
            .Faction = Val(GetWord())
            .Personality(1) = Val(GetWord())
            .Personality(2) = Val(GetWord())
            .AI_Behavior = GetWord()
            .AI_Target = Val(GetWord())
            .reserved = Val(GetWord())
            
            .csvName = .strName  'default
            
            For m = 1 To 3
               .InitPos(m).X = GetWord()
               .InitPos(m).Y = GetWord()
            Next m
            
            .UnknownStr = GetWord()
            
            '��Ա����
            .StacksCount = Val(GetWord())
            If .StacksCount > 0 Then
               ReDim .Stacks(1 To .StacksCount)
               
               For m = 1 To .StacksCount
                  .Stacks(m).ID = Val(GetWord())
                  .Stacks(m).Min = Val(GetWord())
                  .Stacks(m).Max = Val(GetWord())
                  .Stacks(m).Flags = Val(GetWord())
               Next m
            Else
               ReDim .Stacks(0)
            End If
            
            .Degree = GetWord()
            
            .Edit = CheckEditable(EditInfo_PartiesCount, n)
            
            AddIndex .ID, .strID
        End With
        
    Next n
    
    Close #lngHandle
    Pointer = 1

    
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadPartyFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����ReadPartyLine
'**��    �룺Text(String),OutputParty(Type_Party)
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-09 16:44:46
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub ReadPartyLine(Text As String, OutputParty As Type_Party)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
        
    Dim tmp1 As String
    Dim n As Long
    Dim m As Integer
    Dim head As String
    
    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
        With OutputParty
            .UnknownTitle = Val(GetWordL())
            .ID = Val(GetWordL())
            .id2 = Val(GetWordL())
            .strID = GetWordL()
            .strName = GetWordL()
            .Flags = GetWordL()
            .Menu = GetWordL()
            
            .Template = Val(GetWordL())
            .Template_strID = PTs(.Template).ptID
            
            .Faction = Val(GetWordL())
            .Faction_strID = Factions(.Faction).strID
            
            .Personality(1) = Val(GetWordL())
            .Personality(2) = Val(GetWordL())
            .AI_Behavior = GetWordL()
            
            .AI_Target = Val(GetWordL())
            .AI_Target_strID = Parties(.AI_Target).strID
            
            .reserved = Val(GetWordL())
            
            For m = 1 To 3
               .InitPos(m).X = GetWordL()
               .InitPos(m).Y = GetWordL()
            Next m
            
            .UnknownStr = GetWordL()
            
            '��Ա����
            .StacksCount = Val(GetWordL())
            If .StacksCount > 0 Then
               ReDim .Stacks(1 To .StacksCount)
               
               For m = 1 To .StacksCount
                  .Stacks(m).ID = Val(GetWordL())
                  .Stacks(m).strID = Trps(.Stacks(m).ID).strID
                  .Stacks(m).Min = Val(GetWordL())
                  .Stacks(m).Max = Val(GetWordL())
                  .Stacks(m).Flags = Val(GetWordL())
               Next m
            Else
               ReDim .Stacks(0)
            End If
            
            .Degree = GetWordL()

        End With

    LinePointer = 1
    
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ReadPartyLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����LoadMapIconFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-11 00:20:26
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub LoadMapIconFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer, i As Integer, H As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    For n = 0 To 2
        MapIconVersionInform(n) = GetWord()
    Next n
    N_MapIcon = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_MapIconsCount) = N_MapIcon
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_MapIconsCount, CStr(N_MapIcon)
    End If
    
    ReDim MapIcons(N_MapIcon - 1)

    DoEvents
    For n = 0 To N_MapIcon - 1

        With MapIcons(n)
            .ID = n
            
            .strID = GetWord()
            .Flags = Val(GetWord())
            .MeshName = GetWord()

            .mScale = GetWord()
            .Sound = Val(GetWord())
          For m = 0 To 2
            .Offset(m) = GetWord()
          Next m
          
            '����������
       .TriggerCount = Val(GetWord())
       If .TriggerCount > 0 Then
            ReDim Preserve .Triggers(1 To .TriggerCount)
            For i = 1 To .TriggerCount
                  .Triggers(i).tiOn = Val(GetWord())
                  .Triggers(i).ActNum = Val(GetWord())
                  If .Triggers(i).ActNum > 0 Then
                  ReDim Preserve .Triggers(i).tiAct(1 To .Triggers(i).ActNum)
                  For m = 1 To .Triggers(i).ActNum
                       .Triggers(i).tiAct(m).Op = GetWord()
                       .Triggers(i).tiAct(m).ParaNum = Val(GetWord())
                       If .Triggers(i).tiAct(m).ParaNum > 0 Then
                       ReDim Preserve .Triggers(i).tiAct(m).Para(1 To .Triggers(i).tiAct(m).ParaNum)
                       For H = 1 To .Triggers(i).tiAct(m).ParaNum
                            .Triggers(i).tiAct(m).Para(H).Value = GetWord()
                       Next H
                       End If
                  Next m
                  End If
            Next i
       Else
         ReDim .Triggers(0)
       End If
        
            .Edit = CheckEditable(EditInfo_MapIconsCount, n)
            
            AddIndex .ID, .strID
        End With
    Next n
    Close #lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadMapIconFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����ReadMapIconLine
'**��    �룺Text(String),OutputMapIcon(Type_MapIcon)
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-09 16:59:36
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub ReadMapIconLine(Text As String, OutputMapIcon As Type_MapIcon)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim n As Long
    Dim m As Integer, i As Integer, H As Integer
    Dim head As String
    
    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
         With OutputMapIcon
            .strID = GetWordL()
            .Flags = Val(GetWordL())
            .MeshName = GetWordL()

            .mScale = GetWordL()
            
            .Sound = Val(GetWordL())
            .Sound_sndName = Sounds(.Sound).sndName
            
          For m = 0 To 2
            .Offset(m) = GetWordL()
          Next m
          
            '����������
       .TriggerCount = Val(GetWordL())
       If .TriggerCount > 0 Then
            ReDim Preserve .Triggers(1 To .TriggerCount)
            For i = 1 To .TriggerCount
                  .Triggers(i).tiOn = Val(GetWordL())
                  .Triggers(i).ActNum = Val(GetWordL())
                  If .Triggers(i).ActNum > 0 Then
                  ReDim Preserve .Triggers(i).tiAct(1 To .Triggers(i).ActNum)
                  For m = 1 To .Triggers(i).ActNum
                       .Triggers(i).tiAct(m).Op = GetWordL()
                       .Triggers(i).tiAct(m).ParaNum = Val(GetWordL())
                       If .Triggers(i).tiAct(m).ParaNum > 0 Then
                       ReDim Preserve .Triggers(i).tiAct(m).Para(1 To .Triggers(i).tiAct(m).ParaNum)
                       For H = 1 To .Triggers(i).tiAct(m).ParaNum
                            .Triggers(i).tiAct(m).Para(H).Value = GetWordL()
                            BuildQuote_ParamCode .Triggers(i).tiAct(m).Para(H)
                       Next H
                       End If
                  Next m
                  End If
            Next i
       End If
         End With

    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ReadMapIconLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����LoadSoundFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-11 09:06:30
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub LoadSoundFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer, i As Integer, H As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    For n = 0 To 2
        SoundVersionInform(n) = GetWord()
    Next n
    
    'SoundResources
    N_SoundRes = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_SoundRessCount) = N_SoundRes
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_SoundRessCount, CStr(N_SoundRes)
    End If
    
    ReDim SoundRess(N_SoundRes - 1)

    DoEvents
    For n = 0 To N_SoundRes - 1

        With SoundRess(n)
            .ID = n
            
            .sndName = GetWord()
            .Flags = GetWord()
               
            .Edit = CheckEditable(EditInfo_SoundRessCount, n)
            
            AddIndex .ID, .sndName
        End With
    Next n
    
    
    'Sounds
    N_Sound = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_SoundsCount) = N_Sound
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_SoundsCount, CStr(N_Sound)
    End If
    
    ReDim Sounds(N_Sound - 1)

    DoEvents
    For n = 0 To N_Sound - 1

        With Sounds(n)
            .ID = n
            
            .sndName = GetWord()
            .Flags = GetWord()
            
            .ResourceCount = Val(GetWord())
            
            If .ResourceCount > 0 Then
                ReDim .Resource(1 To .ResourceCount)
                
                For m = 1 To .ResourceCount
                     .Resource(m).ID = Val(GetWord())
                     .Resource(m).Unknown = Val(GetWord())
                Next m
            Else
                ReDim .Resource(0)
            End If
               
            .Edit = CheckEditable(EditInfo_SoundsCount, n)
            
            AddIndex .ID, .sndName
        End With
    Next n
    
    Close #lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadSoundFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����ReadSoundLine
'**��    �룺Text(String),OutputSound(Type_Sound)
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-09 17:07:02
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub ReadSoundLine(Text As String, OutputSound As Type_Sound)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer
    
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()

        With OutputSound
        
            .sndName = GetWordL()
            .Flags = GetWordL()
            
            .ResourceCount = Val(GetWordL())
            
            If .ResourceCount > 0 Then
                ReDim .Resource(1 To .ResourceCount)
                
                For m = 1 To .ResourceCount
                     .Resource(m).ID = Val(GetWordL())
                     .Resource(m).strID = SoundRess(.Resource(m).ID).sndName
                     .Resource(m).Unknown = Val(GetWordL())
                Next m
            Else
                ReDim .Resource(0)
            End If
               
        End With
    
    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ReadSoundLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����ReadSoundResLine
'**��    �룺Text(String),OutputSoundRes(Type_Sound)
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-09 17:11:36
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub ReadSoundResLine(Text As String, OutputSoundRes As Type_SoundResource)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer
    
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()

        With OutputSoundRes
            .sndName = GetWordL()
            .Flags = GetWordL()
               
        End With
    
    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ReadSoundResLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����LoadItemFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:24:01
'**�� �� �ˣ�Ser_Charles
'**��    �ڣ�2010-12-08 22:42:16
'**��    ����V1.1321
'*************************************************************************
Sub LoadItemFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer, i As Integer, H As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    For n = 0 To 2
        ItmVersionInform(n) = GetWord()
    Next n
    N_Item = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_ItemsCount) = N_Item
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_ItemsCount, CStr(N_Item)
    End If
    
    ReDim itm(N_Item - 1)

    DoEvents
    For n = 0 To N_Item - 1

        With itm(n)
            .ID = n
            .dbName = GetWord()
            .disname = GetWord()
            .texname = GetWord()
            .nmdl = Val(GetWord())
            
            .csvName = .disname
            .csvName_pl = .disname
            If .nmdl > 0 Then
               ReDim .mdlname(1 To .nmdl)
               ReDim .mdl_b(1 To .nmdl)
               
               For m = 1 To .nmdl
                .mdlname(m) = GetWord()
                .mdl_b(m) = GetWord()
               Next m
            Else
               ReDim .mdlname(0)
               ReDim .mdl_b(0)
            End If

            .itmType = GetWord()
            .Action = GetWord()
            .price = Val(GetWord())
            '.prefix = Val(GetWord())
            .Prefix = GetWord()
            '.weight = Val(GetWord())   'v951
            .weight = GetWord()    'v952
            .abundance = Val(GetWord())
            .head_armor = Val(GetWord())
            .body_armor = Val(GetWord())
            .leg_armor = Val(GetWord())
            .difficulty = Val(GetWord())
            .hit_points = Val(GetWord())
            .speed_rating = Val(GetWord())
            .missile_speed = Val(GetWord())
            .weapon_length = Val(GetWord())
            .max_ammo = Val(GetWord())
            .thrust_damage = Val(GetWord())
            .swing_damage = Val(GetWord())
            
            .FactionCount = Val(GetWord())
            If .FactionCount > 0 Then
               ReDim .Faction(1 To .FactionCount)
               
               For m = 1 To .FactionCount
                .Faction(m).ID = GetWord()
               Next m
            Else
               ReDim .Faction(0)
            End If
            
            
            '����������
       .TriggerCount = Val(GetWord())
       If .TriggerCount > 0 Then
            ReDim .Trigger(1 To .TriggerCount)
            For i = 1 To .TriggerCount
                  .Trigger(i).tiOn = Val(GetWord())
                  .Trigger(i).ActNum = Val(GetWord())
                  If .Trigger(i).ActNum > 0 Then
                     ReDim .Trigger(i).tiAct(1 To .Trigger(i).ActNum)
                     For m = 1 To .Trigger(i).ActNum
                          .Trigger(i).tiAct(m).Op = GetWord()
                          .Trigger(i).tiAct(m).ParaNum = Val(GetWord())
                          If .Trigger(i).tiAct(m).ParaNum > 0 Then
                             ReDim .Trigger(i).tiAct(m).Para(1 To .Trigger(i).tiAct(m).ParaNum)
                             For H = 1 To .Trigger(i).tiAct(m).ParaNum
                                  .Trigger(i).tiAct(m).Para(H).Value = GetWord()
                             Next H
                          Else
                             ReDim .Trigger(i).tiAct(m).Para(0)
                          End If
                     Next m
                  Else
                     ReDim .Trigger(i).tiAct(0)
                  End If
            Next i
       ElseIf .TriggerCount <= 0 Then
            ReDim .Trigger(0)
       End If
        
            .Edit = CheckEditable(EditInfo_ItemsCount, n)
            
            AddIndex .ID, .dbName
        End With
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadItemFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����LoadGlobalVarFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2012-03-02 20:05:43
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub LoadGlobalVarFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim i As Integer, s As String
    Dim F As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    F = FreeFile
    Open tmpFileName For Input As #F
       i = 0
       ReDim gVars(0)
       Do While Not EOF(F)
          Line Input #F, s
          s = Replace(s, "$", "")
          If s <> "" Then
             ReDim Preserve gVars(i)
             gVars(i).VarName = s
             gVars(i).ID = i
             i = i + 1
          End If
       Loop
       N_gVar = i
    Close F

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadGlobalVarFile [n=" & CStr(i) & "]", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����LoadQuickStringFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2012-02-03 22:45:14
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub LoadQuickStringFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    N_qStr = Val(GetWord())
    
    'If MnBInfo.FirstTimeEdit Then
    '    MnBInfo.EditInfo(EditInfo_QuickStringsCount) = N_qStr
    '    WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_QuickStringsCount, CStr(N_qStr)
    'End If
    
    ReDim qStrs(N_qStr - 1)

    DoEvents
    For n = 0 To N_qStr - 1

        With qStrs(n)
            .ID = n
            .Name = GetWord()
            .Str = GetWord()
        
            '.Edit = CheckEditable(EditInfo_QuickStringsCount, n)
            
            'AddIndex .ID, .Name
        End With
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadQuickStringFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����LoadStringFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2012-03-02 22:42:05
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub LoadStringFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim strTem As String
    Dim spcPos As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Input Access Read As lngHandle Len = 1

    Pointer = 1
    
        Line Input #lngHandle, strTem
        StringVersionInform(n) = strTem
        
        Line Input #lngHandle, strTem
        N_Str = Val(strTem)
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_StringsCount) = N_Str
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_StringsCount, CStr(N_Str)
    End If
    
    
    ReDim Strs(N_Str - 1)

    DoEvents
    For n = 0 To N_Str - 1

        Line Input #lngHandle, strTem
        spcPos = InStr(1, strTem, " ")
        If Trim(strTem) <> "" Then
          If spcPos > 0 Then
            With Strs(n)
              .ID = n
              .Name = Left(strTem, spcPos - 1)
              .Str = Right(strTem, Len(strTem) - spcPos)
        
              .Edit = CheckEditable(EditInfo_StringsCount, n)
            
              AddIndex .ID, .Name
            End With
          End If
        End If
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadStringFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����ReadItemLine
'**��    �룺Text(String),(Type_Item)OutputItem
'**��    ����-
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-09 16:12:42
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub ReadItemLine(Text As String, OutputItem As Type_Item, Optional lPointer As Long = 1)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer, i As Integer, H As Integer
    Dim head As String
    
    LinePointer = lPointer
    
    If LinePointer = 1 Then
       txtLine = "h " & Text
       head = GetWordL()
    End If
    
        With OutputItem
            .dbName = GetWordL()
            .disname = GetWordL()
            .texname = GetWordL()
            .nmdl = Val(GetWordL())

            If .nmdl > 0 Then
               ReDim .mdlname(1 To .nmdl)
               ReDim .mdl_b(1 To .nmdl)
               
               For m = 1 To .nmdl
                .mdlname(m) = GetWordL()
                .mdl_b(m) = GetWordL()
               Next m
            Else
               ReDim .mdlname(0)
               ReDim .mdl_b(0)
            End If

            .itmType = GetWordL()
            .Action = GetWordL()
            .price = Val(GetWordL())
            '.prefix = Val(GetWordL())
            .Prefix = GetWordL()
            '.weight = Val(GetWordL())   'v951
            .weight = GetWordL()    'v952
            .abundance = Val(GetWordL())
            .head_armor = Val(GetWordL())
            .body_armor = Val(GetWordL())
            .leg_armor = Val(GetWordL())
            .difficulty = Val(GetWordL())
            .hit_points = Val(GetWordL())
            .speed_rating = Val(GetWordL())
            .missile_speed = Val(GetWordL())
            .weapon_length = Val(GetWordL())
            .max_ammo = Val(GetWordL())
            .thrust_damage = Val(GetWordL())
            .swing_damage = Val(GetWordL())
            
            .FactionCount = Val(GetWordL())
            If .FactionCount > 0 Then
               ReDim .Faction(1 To .FactionCount)
               
               For m = 1 To .FactionCount
                .Faction(m).ID = GetWordL()
                .Faction(m).strID = Factions(.Faction(m).ID).strID
               Next m
            Else
               ReDim .Faction(0)
            End If
            
            
            '����������
       .TriggerCount = Val(GetWordL())
       If .TriggerCount > 0 Then
            ReDim .Trigger(1 To .TriggerCount)
            For i = 1 To .TriggerCount
                  .Trigger(i).tiOn = Val(GetWordL())
                  .Trigger(i).ActNum = Val(GetWordL())
                  If .Trigger(i).ActNum > 0 Then
                     ReDim .Trigger(i).tiAct(1 To .Trigger(i).ActNum)
                     For m = 1 To .Trigger(i).ActNum
                          .Trigger(i).tiAct(m).Op = GetWordL()
                          .Trigger(i).tiAct(m).ParaNum = Val(GetWordL())
                          If .Trigger(i).tiAct(m).ParaNum > 0 Then
                             ReDim .Trigger(i).tiAct(m).Para(1 To .Trigger(i).tiAct(m).ParaNum)
                             For H = 1 To .Trigger(i).tiAct(m).ParaNum
                                 .Trigger(i).tiAct(m).Para(H).Value = GetWordL()
                                  BuildQuote_ParamCode .Trigger(i).tiAct(m).Para(H)
                             Next H
                          Else
                             ReDim .Trigger(i).tiAct(m).Para(0)
                          End If
                     Next m
                  Else
                     ReDim .Trigger(i).tiAct(0)
                  End If
            Next i
       ElseIf .TriggerCount <= 0 Then
            ReDim .Trigger(0)
       End If
        
        End With

    lPointer = LinePointer

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ReadItemLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SaveItemFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:24:04
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-11-30 22:59:07
'**��    ����V0.951.12
'*************************************************************************
Sub SaveItemFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim n As Long
    Dim m As Integer

    Dim txtLine As String

    Dim lngFileHandle As Integer
   
    lngFileHandle = FreeFile()
    Open FilePath For Output As #lngFileHandle

    txtLine = ""

    For n = 0 To 1
        txtLine = txtLine & ItmVersionInform(n) & " "
    Next n

    txtLine = txtLine & ItmVersionInform(2) & vbCrLf

    txtLine = txtLine & CStr(N_Item)
    Print #lngFileHandle, txtLine

    For n = 0 To N_Item - 1
            txtLine = OutputItemLine(n)
            Print #lngFileHandle, txtLine
    Next n

    Close #lngFileHandle

    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveItemFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SaveTriggerFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-02-23 14:11:36
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Sub SaveTriggerFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim n As Long
    Dim m As Integer

    Dim txtLine As String

    Dim lngFileHandle As Integer
   
    lngFileHandle = FreeFile()
    Open FilePath For Output As #lngFileHandle

    txtLine = ""

    For n = 0 To 1
        txtLine = txtLine & TimeTrgVersionInform(n) & " "
    Next n

    txtLine = txtLine & TimeTrgVersionInform(2) & vbCrLf

    txtLine = txtLine & CStr(N_TimeTrg)
    Print #lngFileHandle, txtLine

    For n = 0 To N_TimeTrg - 1
            txtLine = OutputTimeTriggerLine(n)
            Print #lngFileHandle, txtLine
    Next n

    Close #lngFileHandle

    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveTriggerFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SaveStringFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2012-03-03 23:39:36
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Sub SaveStringFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim n As Long
    Dim m As Integer

    Dim txtLine As String

    Dim lngFileHandle As Integer
   
    lngFileHandle = FreeFile()
    Open FilePath For Output As #lngFileHandle

    txtLine = ""
    
    txtLine = txtLine & StringVersionInform(0) & vbCrLf

    txtLine = txtLine & CStr(N_Str)
    Print #lngFileHandle, txtLine

    For n = 0 To N_Str - 1
            'Strs(n).Name = Replace(Strs(n).Name, " ", "_")
            'Strs(n).Str = Replace(Strs(n).Str, " ", "_")
            txtLine = Strs(n).Name & " " & Strs(n).Str
            Print #lngFileHandle, txtLine
    Next n

    Close #lngFileHandle

    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveStringFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SavePSysFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-01-07 16:26:23
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Sub SavePSysFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim n As Long
    Dim m As Integer

    Dim txtLine As String

    Dim lngFileHandle As Integer
   
    lngFileHandle = FreeFile()
    Open FilePath For Output As #lngFileHandle

    txtLine = ""

    For n = 0 To 1
        txtLine = txtLine & PSysVersionInform(n) & " "
    Next n

    txtLine = txtLine & PSysVersionInform(2) & vbCrLf

    txtLine = txtLine & CStr(N_PSys)
    Print #lngFileHandle, txtLine

    For n = 0 To N_PSys - 1
            txtLine = OutputPSysLine(n)
            Print #lngFileHandle, txtLine
    Next n

    Close #lngFileHandle

    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SavepsysFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SaveTabMatFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-01-22 19:46:54
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Sub SaveTabMatFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim n As Long
    Dim m As Integer

    Dim txtLine As String

    Dim lngFileHandle As Integer
   
    lngFileHandle = FreeFile()
    Open FilePath For Output As #lngFileHandle

    txtLine = ""

    txtLine = txtLine & CStr(N_TabMat)
    Print #lngFileHandle, txtLine

    For n = 0 To N_TabMat - 1
            txtLine = OutputTabMatLine(n)
            Print #lngFileHandle, txtLine
    Next n

    Close #lngFileHandle

    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveTabMatFile", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����SaveMeshFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-01-07 16:26:23
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Sub SaveMeshFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim n As Long
    Dim m As Integer

    Dim txtLine As String

    Dim lngFileHandle As Integer
   
    lngFileHandle = FreeFile()
    Open FilePath For Output As #lngFileHandle

    txtLine = ""

    'For n = 0 To 1
    '    txtLine = txtLine & MeshVersionInform(n) & " "
    'Next n

    'txtLine = txtLine & MeshVersionInform(2) & vbCrLf

    txtLine = txtLine & CStr(N_Mesh)
    Print #lngFileHandle, txtLine

    For n = 0 To N_Mesh - 1
            txtLine = OutputMeshLine(n)
            Print #lngFileHandle, txtLine
    Next n

    Close #lngFileHandle

    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveMeshFile", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����InitTags
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-02-07 11:05:54
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.132.1
'*************************************************************************
Sub InitTags()
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Tags(Tag_Register) = "reg"  '�Զ���tag
    Tags(Tag_Variable) = "var"  '�Զ���tag
    Tags(Tag_String) = "str"
    Tags(Tag_Item) = "itm"
    Tags(Tag_Troop) = "trp"
    Tags(Tag_Faction) = "fac"
    Tags(Tag_Quest) = "qst"
    Tags(Tag_Party_Tpl) = "pt"
    Tags(Tag_Party) = "p"
    Tags(Tag_Scene) = "scn"
    Tags(Tag_Mission_tpl) = "mst"
    Tags(Tag_Menu) = "menu"
    Tags(Tag_Script) = "script"   '�Զ���tag
    Tags(Tag_Particle_Sys) = "psys"
    Tags(Tag_Scene_Prop) = "spr"
    Tags(Tag_Sound) = "snd"
    Tags(Tag_Local_Variable) = "lvar"  '�Զ���tag
    Tags(Tag_Map_Icon) = "icon"     '�Զ���tag
    Tags(Tag_Skill) = "skl"
    Tags(Tag_Mesh) = "mesh"
    Tags(Tag_Presentation) = "prsnt"
    Tags(Tag_Quick_String) = "qstr"
    Tags(Tag_Track) = "track"     '�Զ���tag
    Tags(Tag_Tableau) = "tab"
    Tags(Tag_Animation) = "anim"   '�Զ���tag
    Tags(Tags_End) = "end"

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "InitTags", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����InitShortTags
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-02-07 11:05:54
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.132.1
'*************************************************************************
Sub InitShortTags()
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    ShortTags(Tag_Register).X = 720
    ShortTags(Tag_Variable).X = 1441
    ShortTags(Tag_String).X = 2161
    ShortTags(Tag_Item).X = 2882
    ShortTags(Tag_Troop).X = 3602
    ShortTags(Tag_Faction).X = 4323
    ShortTags(Tag_Quest).X = 5044
    ShortTags(Tag_Party_Tpl).X = 5764
    ShortTags(Tag_Party).X = 6485
    ShortTags(Tag_Scene).X = 7205
    ShortTags(Tag_Mission_tpl).X = 7926
    ShortTags(Tag_Menu).X = 8646
    ShortTags(Tag_Script).X = 9367
    ShortTags(Tag_Particle_Sys).X = 10088
    ShortTags(Tag_Scene_Prop).X = 10808
    ShortTags(Tag_Sound).X = 11529
    ShortTags(Tag_Local_Variable).X = 12249
    ShortTags(Tag_Map_Icon).X = 12970
    ShortTags(Tag_Skill).X = 13690
    ShortTags(Tag_Mesh).X = 14411
    ShortTags(Tag_Presentation).X = 15132
    ShortTags(Tag_Quick_String).X = 15852
    ShortTags(Tag_Track).X = 16573
    ShortTags(Tag_Tableau).X = 17293
    ShortTags(Tag_Animation).X = 18014
    ShortTags(Tags_End).X = 18734

    ShortTags(Tag_Register).Y = 7927936
    ShortTags(Tag_Variable).Y = 5855872
    ShortTags(Tag_String).Y = 3783808
    ShortTags(Tag_Item).Y = 1711744
    ShortTags(Tag_Troop).Y = 9639680
    ShortTags(Tag_Faction).Y = 7567616
    ShortTags(Tag_Quest).Y = 5495552
    ShortTags(Tag_Party_Tpl).Y = 3423488
    ShortTags(Tag_Party).Y = 1351424
    ShortTags(Tag_Scene).Y = 9279360
    ShortTags(Tag_Mission_tpl).Y = 7207296
    ShortTags(Tag_Menu).Y = 5135232
    ShortTags(Tag_Script).Y = 3063168
    ShortTags(Tag_Particle_Sys).Y = 991104
    ShortTags(Tag_Scene_Prop).Y = 8919040
    ShortTags(Tag_Sound).Y = 6846976
    ShortTags(Tag_Local_Variable).Y = 4774912
    ShortTags(Tag_Map_Icon).Y = 2702848
    ShortTags(Tag_Skill).Y = 630784
    ShortTags(Tag_Mesh).Y = 8558720
    ShortTags(Tag_Presentation).Y = 6486656
    ShortTags(Tag_Quick_String).Y = 4414592
    ShortTags(Tag_Track).Y = 2342528
    ShortTags(Tag_Tableau).Y = 270464
    ShortTags(Tag_Animation).Y = 8198400
    ShortTags(Tags_End).Y = 6126336
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "InitShortshorttags", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����inititmpicker
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-06-14 06:43:27
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.955.6
'*************************************************************************
Sub InitItmPicker()
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim iii As Long
    Dim Scrollwidth As Long
    Dim res As Long
    Dim strTmp As String

    Scrollwidth = 0

    'Form1.itmpicker.Clear
    For iii = 0 To N_Item - 1
        If Len(itm(iii).csvName) > 0 Then
            strTmp = CStr(iii) & " " & itm(iii).csvName & " " & itm(iii).disname
        Else
            strTmp = CStr(iii) & " " & itm(iii).disname & " " & itm(iii).csvName
        End If

        If Len(strTmp) > Scrollwidth Then
            Scrollwidth = Len(strTmp)
        End If
        'Form1.itmpicker.AddItem (strTmp)
    Next iii

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "InitItmPicker", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����initItmTypeList
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:24:14
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Sub initItmTypeList(objListBox As ListBox)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    objListBox.Clear

    Call ItmTypeListAddItem(objListBox, 0, "ȫ��")
    Call ItmTypeListAddItem(objListBox, itp_type_horse, "��ƥ")
    Call ItmTypeListAddItem(objListBox, itp_type_one_handed_wpn, "��������")
    Call ItmTypeListAddItem(objListBox, itp_type_two_handed_wpn, "˫������")
    Call ItmTypeListAddItem(objListBox, itp_type_polearm, "��������")
    Call ItmTypeListAddItem(objListBox, itp_type_arrows, "��")
    Call ItmTypeListAddItem(objListBox, itp_type_bolts, "���")
    Call ItmTypeListAddItem(objListBox, itp_type_shield, "����")
    Call ItmTypeListAddItem(objListBox, itp_type_bow, "��")
    Call ItmTypeListAddItem(objListBox, itp_type_crossbow, "��")
    Call ItmTypeListAddItem(objListBox, itp_type_thrown, "Ͷ������")
    Call ItmTypeListAddItem(objListBox, itp_type_goods, "����")
    Call ItmTypeListAddItem(objListBox, itp_type_head_armor, "ͷ��")
    Call ItmTypeListAddItem(objListBox, itp_type_body_armor, "����")
    Call ItmTypeListAddItem(objListBox, itp_type_foot_armor, "Ь��")
    Call ItmTypeListAddItem(objListBox, itp_type_hand_armor, "����")
    Call ItmTypeListAddItem(objListBox, itp_type_pistol, "��ǹ")
    Call ItmTypeListAddItem(objListBox, itp_type_musket, "��ǹ")
    Call ItmTypeListAddItem(objListBox, itp_type_bullets, "�ӵ�")
    Call ItmTypeListAddItem(objListBox, itp_type_animal, "����")
    Call ItmTypeListAddItem(objListBox, itp_type_book, "�鼮")

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "initItmTypeList", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����ItmTypeListAddItem
'**��    �룺itp_type_code(Byte)   -
'**        ��itp_type_name(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:24:17
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Sub ItmTypeListAddItem(objListBox As ListBox, itp_type_code As Byte, itp_type_name As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim tmpStr As String
    tmpStr = TranslateStr("Form3_typeopt", CStr(itp_type_code), itp_type_name)
    objListBox.AddItem (CStr(itp_type_code) & " " & tmpStr)

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ItmTypeListAddItem", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����showHexId
'**��    �룺(String)id
'**��    ����(String)
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:24:30
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-12-13 11:11:15
'**��    ����V1.1321
'*************************************************************************
Function showHexId(ID As String) As String
    On Error Resume Next
    showHexId = CStr(Hex(Val(ID)))
End Function


'*************************************************************************
'**�� �� ����PublicInit
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-17 18:42:18
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-11-23 22:40:09
'**��    ����V1.1321
'*************************************************************************
Sub PublicInit()
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    CurrentTrpID = -1
    CurPartyTemplateID = -1
    CurrentItmID = -1
    CurrentPartyID = -1
    CurrentSceneID = -1
    CurrentFactionID = -1
    CurrentMapIconID = -1
    CurrentPSysID = -1
    
    ReDim Trps(0)
    ReDim itm(0)
    ReDim Scenes(0)
    ReDim Parties(0)
    ReDim PTs(0)
    ReDim Factions(0)
    ReDim MapIcons(0)
    ReDim PSys(0)
    
    'Dim i As Integer
    
    'For i = 0 To 99
    '    DreamTeam(i) = -1
    'Next i
    
    Init_Integer64b
    
    RunERR.MissCSV = False
    RunERR.MissINI = False
    RunERR.MissMod = False
        
    'frmData.Show
    InitTags
    InitPYTags
    InitShortTags
    InitPy
    
    InitPublicWords
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "PublicInit", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����GetTroopSkill
'**��    �룺TroopID(integer,-1��Ϊ��ǰ����), ByVal sklid(Integer) -
'**��    ����(Byte) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-04-17 20:06:12
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Function GetTroopSkill(TroopID As Integer, ByVal sklid As Integer) As Byte
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim a As Integer64b, s As String, b As String
    
    sklid = sklid - 1
    
    If TroopID = -1 Then TroopID = CurrentTrpID
    If TroopID = -1 Then Exit Function
    
    With Trps(TroopID)
        
        a = StrToI64(.Skills((sklid \ 8) + 1))
        s = I64toHexStr(a)
        s = Mid(s, Len(s) - ((sklid) Mod 8), 1)
    
    End With
    GetTroopSkill = HexStrToI4(s)
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "GetTroopSkill", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����GetSkill
'**��    �룺ByVal sklid(Integer) -
'**��    ����(Byte) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:24:38
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************
Function GetSkill(ByVal sklid As Integer) As Byte
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim a As Integer64b, s As String, b As String
    
    sklid = sklid - 1
    
    If CurrentTrpID = -1 Then Exit Function
    
    With Trps(CurrentTrpID)
        
        a = StrToI64(.Skills((sklid \ 8) + 1))
        s = I64toHexStr(a)
        s = Mid(s, Len(s) - ((sklid) Mod 8), 1)
    
    End With
    GetSkill = HexStrToI4(s)
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "GetSkill", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����PutSkill
'**��    �룺intVal(Integer)         -
'**        ��ByVal sklid(Integer) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:24:41
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-11-24 15:23:09
'**��    ����V1.1321
'*************************************************************************
Sub PutSkill(Trp As Type_Troops, intVal As Integer, ByVal sklid As Integer)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim a As Integer64b, s As String, b As String
    Dim n As Long
    
    Call Init_Integer64b

    With Trp
    sklid = sklid - 1
    a = StrToI64(.Skills((sklid \ 8) + 1))
    s = I64toHexStr(a)
        For n = 0 To 7
            If n <> (sklid Mod 8) Then
                b = Mid(s, Len(s) - n, 1) + b
            Else
                b = CHex(intVal) + b
            End If
        Next n
        .Skills((sklid \ 8) + 1) = I64toStrNZ(HexStrToI64(b))
    End With
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "PutSkill", Err.Number, Err.Description)
End Sub

Sub load_KIS_Team()
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim troopsID As Long
    Dim tempStr As String
    Dim n As Integer
    Dim strTmp As String
    Dim intLoop As Integer

    CountDream = 0

    strTmp = ReadString(MnBInfo.ModIniFileName, "KIS_Team", "GuardsManCount", 250)

    If Val(strTmp) > 0 Then
        intLoop = Val(strTmp)
    Else
        intLoop = 99
    End If
    For n = 0 To Val(strTmp)
        strTmp = ReadString(MnBInfo.ModIniFileName, "KIS_Team", "GuardsMan_" & CStr(n), 250)

        If Len(strTmp) = 0 Or strTmp = "0" Or strTmp = "-1" Then
            strTmp = "-1"
        Else
            troopsID = Val(strTmp)
            DreamTeam(n) = troopsID

            strTmp = fnNumberFixedLength(CLng(CountDream), 3, " ")
            tempStr = strTmp & "["

            strTmp = fnNumberFixedLength(troopsID, 3, "")
            tempStr = tempStr & strTmp & "]"

            If Len(Trps(troopsID).csvName) = 0 Then
                tempStr = tempStr & " " & Trps(troopsID).strID & " "
            Else
                tempStr = tempStr & " " & Trps(troopsID).csvName & " "
            End If

            'PTForm.Dream_list.AddItem tempStr
            CountDream = CountDream + 1
        End If
    Next n

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("PTForm", "KIS_Team", Err.Number, Err.Description)
End Sub




'*************************************************************************
'**�� �� ����updateTriggersFile
'**��    �룺id(Integer)   - trigger��ID�����µ�id��trigger��0=��ӡ�
'**        ��trigger(String) - trigger���ݡ�
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-06-14 09:43:09
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.955.6
'*************************************************************************
Public Function updateTriggersFile(ID As Integer, Trigger As String) As Integer
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim strFileName As String
    Dim fileLineNum As Integer
    Dim strTemp As String
    Dim fileBuff As String
    Dim Count As Long
    Dim opFlag As Boolean

    opFlag = False

    updateTriggersFile = -1
    
    If ID < 0 Then Exit Function
    If ID = 0 And Len(Trigger) = 0 Then Exit Function

    Count = 0

    strFileName = MnBInfo.ModName & "\triggers.txt"
    MaxPointer = FileLen(strFileName)
    lngHandle = FreeFile()
    Open strFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    fileLineNum = 0
    Do Until 0
        strTemp = GetRealLine()
        'fix me
        If Len(strTemp) = 0 Then
            'If fileLineNum >= count + 2 Then
            '    strTemp = GetRealLine()
            'End If
        End If
        fileLineNum = fileLineNum + 1
        If Pointer > MaxPointer Then
            If ID = 0 Then
                fileBuff = fileBuff & strTemp & vbCrLf
                'append trigger .
                fileBuff = fileBuff & Trigger
            Else
                If opFlag = False Then
                    fileBuff = fileBuff & Trigger
                Else
                    fileBuff = fileBuff & strTemp
                End If
            End If
            Exit Do
        End If

        If fileLineNum = 2 Then
            'get trigger ������¼.
            Count = Val(strTemp)
        End If

        ' add mode
        If ID = 0 Then
            If fileLineNum = 2 Then
                '���� trigger ������¼.
                Count = Count + 1
                fileBuff = fileBuff & CStr(Count) & vbCrLf
            Else
                fileBuff = fileBuff & strTemp & vbCrLf
            End If
        Else
            'update mode
            If Count > 0 And fileLineNum = 2 + ID Then
                '�滻 trigger .
                fileBuff = fileBuff & Trigger & vbCrLf
                opFlag = True
            Else
                fileBuff = fileBuff & strTemp & vbCrLf
            End If

        End If

    Loop
    Close #lngHandle

    'write to file
    lngHandle = FreeFile()
    Open strFileName For Output As #lngHandle
    Print #lngHandle, fileBuff
    Close #lngHandle

    'fix me
    'OutAsDebugTex (fileBuff)

    updateTriggersFile = fileLineNum - 2

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    updateTriggersFile = -1
    Call logErr("ModMain", "updateTriggersFile", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����FixPath
'**��    �룺path(String) -
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-06-21 06:54:18
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-11-21 23:13:54
'**��    ����V1.132.21
'*************************************************************************
Public Function FixPath(Path As String) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    FixPath = Trim$(Path)
    If Len(Path) = 0 Then
        Exit Function
    End If

    Dim i As Long
    If Right(FixPath, 1) = "\" Then
        FixPath = Left(FixPath, Len(FixPath) - 1)
    End If

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "FixPath", Err.Number, Err.Description)
End Function


'*************************************************************************
'**�� �� ����get_Module_Version
'**��    �룺ModIniFilePath(String) -
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����MnBInfo.ModPath
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-06-21 06:43:50
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.955.23
'*************************************************************************
Private Function get_Module_Version(ModIniFilePath As String) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim TemP As String
    Dim arrTmp() As String
    Dim Key As String
    Dim ii As Integer
    
    Dim strKeyFileName As String
    strKeyFileName = "module.ini"

    Key = "works_with_version_max"
    ii = Len(Key)

    lngHandle = FreeFile()

    Open MnBInfo.ModPath & strKeyFileName For Input As #lngHandle
    Do Until EOF(lngHandle)
        Input #lngHandle, TemP
        If LCase$(Trim$(Left$(TemP, ii))) = Key Then
            arrTmp = Split(TemP, "=")
            If LCase$(Trim$(arrTmp(0))) = Key Then
                get_Module_Version = LCase$(Trim$(arrTmp(1)))
                Exit Function
            End If
        End If
    Loop
    Close #lngHandle

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    'Call logErr("ModMain", "get_Module_Version", Err.Number, Err.Description)
    get_Module_Version = "--"
End Function

Public Function computeListID(objListBox As ListBox, intMode As Integer) As Long
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    computeListID = -1
    If objListBox Is Nothing Then Exit Function

    Dim tempStr As String
    Dim arrTmp() As String
    Dim i As Integer

    If intMode = 1 Then
        tempStr = objListBox.Text
        
        arrTmp = Split(tempStr, " ")
        
        For i = 0 To UBound(arrTmp)
            arrTmp(i) = Trim$(arrTmp(i))
            If Len(arrTmp(i)) > 0 Then
                tempStr = arrTmp(i)
                Exit For
            End If
        Next

        computeListID = Val(tempStr)
    Else
        computeListID = objListBox.ListIndex
    End If

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "computeListID", Err.Number, Err.Description)
End Function


'*************************************************************************
'**�� �� ����ShowItemListByType
'**��    �룺objItemListBox(ListBox)     -
'**        ��objItemTypeListBox(ListBox) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-06-29 07:56:16
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.960.7
'*************************************************************************
Public Sub ShowItemListByType(objItemListBox As ListBox, objItemTypeListBox As ListBox)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim itmType As Integer
    Dim intShowItemType As Integer
    Dim tmpCount As Long
    Dim Scrollwidth As Long
    Dim strTmp As String
    
    Dim n As Long

    If objItemListBox Is Nothing Then Exit Sub
    If objItemTypeListBox Is Nothing Then Exit Sub

    If objItemTypeListBox.ListIndex = -1 Then
        Exit Sub
    End If

    intShowItemType = objItemTypeListBox.ListIndex

    tmpCount = 0
    ReDim itmID(0)
    objItemListBox.Clear
    Scrollwidth = 0

    For n = 0 To N_Item - 1

        If intShowItemType = 0 Then
            ' show all
            itmType = 0
        Else
            If Val(itm(n).itmType) < 2100000000 Then
                itmType = Val(itm(n).itmType) Mod 256
            Else
                itmType = MinusIT(itm(n).itmType) Mod 256
            End If
        End If

        If itmType = intShowItemType Then
            If Len(itm(n).csvName) > 0 Then
                strTmp = CStr(n) & " " & itm(n).csvName & " " & itm(n).disname
            Else
                strTmp = CStr(n) & " " & itm(n).disname & " " & itm(n).csvName
            End If

            If Len(strTmp) > Scrollwidth Then
                Scrollwidth = Len(strTmp)
            End If

            objItemListBox.AddItem (strTmp)
            itmID(tmpCount) = n
            tmpCount = tmpCount + 1
            ReDim Preserve itmID(tmpCount)
        End If
    Next n

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ShowItemListByType", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**�� �� ����getTXTID
'**��    �룺tag_type(Integer) -
'**        ��id(Long)          -
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-07-27 06:48:49
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.960.38
'*************************************************************************
Public Function getTXTID(Tag_Type As Integer, ID As Long, Optional strID As String) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    'tag_register        =  1
    'tag_variable        =  2
    'tag_string          =  3
    'tag_item            =  4
    'tag_troop           =  5
    'tag_faction         =  6
    'tag_quest           =  7
    'tag_party_tpl       =  8
    'tag_party           =  9
    'tag_scene           = 10
    'tag_mission_tpl     = 11
    'tag_menu            = 12
    'tag_script          = 13
    'tag_particle_sys    = 14
    'tag_scene_prop      = 15
    'tag_sound           = 16
    'tag_local_variable  = 17
    'tag_map_icon        = 18
    'tag_skill           = 19
    'tag_mesh            = 20
    'tag_presentation    = 21
    'tag_quick_string    = 22
    'tag_track           = 23
    'tag_tableau         = 24
    'tag_animation       = 25
    'tags_end            = 26

    Dim i64b_num1 As Integer64b
    Dim i64b_num2 As Integer64b
    
    Call Init_Integer64b

    i64b_num1 = HexStrToI64(CStr(Hex$(Tag_Type)) & "00000000000000")
    i64b_num2 = StrToI64(CStr(ID))
    i64b_num2 = Plus64b(i64b_num1, i64b_num2)
    
    getTXTID = I64toStrNZ(i64b_num2)
    
    If Not IsMissing(strID) Then
        strID = GetstrID(Tag_Type, CStr(ID))
    End If
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "getTXTID", Err.Number, Err.Description)

End Function

'*************************************************************************
'**�� �� ����GetstrID
'**��    �룺Tag_Type(Integer),Index(String) -
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-03-21 20:23:53
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function GetstrID(Tag_Type As Integer, Index As String) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
  Select Case Tag_Type
    Case Tag_Register
    
    Case Tag_Variable
    
    Case Tag_String
    Case Tag_Item
         GetstrID = itm(Val(Index)).dbName
    Case Tag_Troop
         GetstrID = Trps(Val(Index)).strID
    Case Tag_Faction
         GetstrID = Factions(Val(Index)).strID
    Case Tag_Quest
    Case Tag_Party_Tpl
         GetstrID = PTs(Val(Index)).ptID
    Case Tag_Party
         GetstrID = Parties(Val(Index)).strID
    Case Tag_Scene
         GetstrID = Scenes(Val(Index)).strID
    Case Tag_Mission_tpl
    Case Tag_Menu
    Case Tag_Script
    Case Tag_Particle_Sys
         GetstrID = PSys(Val(Index)).strID
    Case Tag_Scene_Prop
    Case Tag_Sound
         GetstrID = Sounds(Val(Index)).sndName
    Case Tag_Local_Variable
    Case Tag_Map_Icon
         GetstrID = MapIcons(Val(Index)).strID
    Case Tag_Skill
    Case Tag_Mesh
         GetstrID = Mesh(Val(Index)).strID
    Case Tag_Presentation
    Case Tag_Quick_String
    Case Tag_Track
    Case Tag_Tableau
         GetstrID = TabMat(Val(Index)).strID
    Case Tag_Animation
    Case Tags_End
  End Select
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "GetstrID", Err.Number, Err.Description)

End Function


'*************************************************************************
'**�� �� ����InitWarbandInfo
'**��    �룺��
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-21 22:52:06
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function InitWarbandInfo(Optional MBHome As String = "", Optional Version As String = "") As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim strTmp As String
    Dim i As Integer, s As Long
    
    InitWarbandInfo = False
    
    DisplayMode = 0
    
    MnBInfo.InfoFinished = False
    MnBInfo.InitFinished = False
    MnBInfo.iniSetting = App.Path & "\" & App.EXEName & ".ini"
    MnBInfo.Language = "en"
    If Trim(MBHome) = "" Then
      strTmp = ReadRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\Mount&Blade Warband", "")
    Else
      strTmp = MBHome
    End If
    
    ' fix me
    strTmp = FixPath(strTmp)
    If Not DirExists(strTmp) Then
         MsgBox "���ĵ�����û�а�װ�������뿳ɱ:ս�š�!", vbCritical, "Error!"
         Exit Function
    End If
With MnBInfo
     .MBHome = strTmp
     .ModPath = strTmp & "\Modules"
     
     If Trim(Version) = "" Then
       .Version = ReadRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\Mount&Blade Warband", "Version")
     Else
       .Version = Version
     End If
     
     strTmp = GetMyDocumentDirectory()
     strTmp = FixPath(strTmp)
     
     
     .MBsaves = strTmp & "\Mount&Blade Warband Savegames"
     .MBsets = strTmp & "\Mount&Blade Warband"
     If FileExists(.MBsets & "\language.txt") Then
            .Language = QuickReadFile(.MBsets & "\language.txt")
     Else
            strTmp = GetAppDataPath
            strTmp = FixPath(strTmp)
            .MBsets = strTmp & "\Mount&Blade Warband"
            .Language = QuickReadFile(.MBsets & "\language.txt")
     End If
     
     .Language = Trim(.Language)
     If .Language = "" Then
         .Language = "en"
     End If
     
     strTmp = Trim(ReadString(.iniSetting, "Settings", "Language", 250))
     SelectLanguage strTmp
     
     LoadOperations
     
     .InitFinished = True
End With
    InitWarbandInfo = True
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "InitWarbandInfo", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����FinishWarbandInfo
'**��    �룺(String)ModName
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-22 22:20:40
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function FinishWarbandInfo(ByVal ModName As String) As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim strTem As String, i As Integer
With MnBInfo
     .ModPath = .MBHome & "\Modules\" & ModName
     .ModBackUp = .ModPath & "\BackUp"
     .ModName = ModName
     .ModIniFileName = .ModPath & "\module.ini"
     .iniFileName = .ModPath & "\MnBWarband_Editor.ini"
     .FirstTimeEdit = False
     
     If FileExists(.iniFileName) Then
       
        For i = 0 To UBound(.EditInfo)
           .EditInfo(i) = ReadInt(.iniFileName, "EDITINFO", "Count" & i)
           
           If .EditInfo(i) = 0 Then
             .FirstTimeEdit = True
             Exit For
           End If
        Next i
       
     Else
       .FirstTimeEdit = True
     End If
     
     LoadModINI
     
     'frmData.Show
     .InfoFinished = True
End With
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "FinishWarbandInfo", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����QuickReadFile
'**��    �룺(String)FilePath
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-22 23:26:26
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function QuickReadFile(ByVal FilePath As String) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim F As Integer, strTem As String
    
    F = FreeFile
    
    Open FilePath For Input As #F
          Line Input #F, strTem
          
          QuickReadFile = strTem
    Close #F
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "QuickReadFile", Err.Number, Err.Description)
End Function

Public Function IsDirectory(ByVal Path As String) As Boolean
If GetAttr(Path) And vbDirectory Then
      If Right(Path, 1) <> "." Then
          IsDirectory = True
      End If
End If
End Function

'*************************************************************************
'**�� �� ����MnBtoRGBColor
'**��    �룺(String)MnBColor
'**��    ����(String) -
'**����������ʵ���￳��ɫ��RGB��ɫ��ת��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-24 22:58:28
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function MnBtoRGBColor(ByVal MnBColor As String) As String
Dim r As Byte, G As Byte, b As Byte, MnBColorI64 As Integer64b, RGBColorI64 As Integer64b
'�ֽ�RGB��ɫֵ
MnBColorI64 = StrToI64(MnBColor)

r = MnBColorI64.by(2)
G = MnBColorI64.by(1)
b = MnBColorI64.by(0)
RGBColorI64.by(0) = r
RGBColorI64.by(1) = G
RGBColorI64.by(2) = b

MnBtoRGBColor = I64toHexStr(RGBColorI64)
MnBtoRGBColor = Right(MnBtoRGBColor, 6)
'RGB-BGR
End Function

'*************************************************************************
'**�� �� ����RGBtoMnBColor
'**��    �룺(String)RGBColor
'**��    ����(String) -
'**����������ʵ���￳��ɫ��RGB��ɫ��ת��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-24 22:58:28
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function RGBtoMnBColor(ByVal RGBColor As String) As String
Dim r As Byte, G As Byte, b As Byte, MnBColorI64 As Integer64b, RGBColorI64 As Integer64b
'�ֽ�RGB��ɫֵ
RGBColorI64 = HexStrToI64(RGBColor)

b = RGBColorI64.by(2)
G = RGBColorI64.by(1)
r = RGBColorI64.by(0)
MnBColorI64.by(0) = b
MnBColorI64.by(1) = G
MnBColorI64.by(2) = r

RGBtoMnBColor = I64toStrNZ(MnBColorI64)
'RGB-BGR
End Function

'*************************************************************************
'**�� �� ����FindItem
'**��    �룺...
'**��    ����(ListItem) -
'**��������������ListView��ѯ����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-27 21:19:17
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function FindItem(oListView As ListView, Start As Long, Row As String, KeyWord As String, bPartial As Boolean, Compare As VbCompareMethod, Optional bReverse As Boolean = False) As ListItem
If oListView.ListItems.Count > 0 And Start <= oListView.ListItems.Count And Start >= 0 And Trim(Row) <> "" Then
Dim i As Long, n As Long, r() As String, j As Long, Finded As Boolean, SearchInfo(2) As Long

If Not bReverse Then
SearchInfo(0) = Start + 1
SearchInfo(1) = oListView.ListItems.Count
SearchInfo(2) = 1
Else
SearchInfo(0) = Start - 1
SearchInfo(1) = 1
SearchInfo(2) = -1
End If

r = Split(Row, "|")
    For i = SearchInfo(0) To SearchInfo(1) Step SearchInfo(2)
    With oListView.ListItems(i)
    
    Finded = False
            For j = 0 To UBound(r)
               If Val(r(j)) = 0 Then
                  n = InStr(1, .Text, KeyWord, Compare)
                 If bPartial Then
                   If n > 0 Then Finded = True
                 Else
                   If n = 1 And Len(KeyWord) = Len(.Text) Then Finded = True
                 End If
                  If Finded Then
                    Set FindItem = oListView.ListItems(i)
                    Exit Function
                  End If
            
               Else
                  n = InStr(1, .SubItems(Val(r(j))), KeyWord, Compare)
                    If bPartial Then
                       If n > 0 Then Finded = True
                    Else
                       If n = 1 And Len(KeyWord) = Len(.SubItems(Val(r(j)))) Then Finded = True
                    End If
                  If Finded Then
                    Set FindItem = oListView.ListItems(i)
                    Exit Function
                  End If
               End If
            Next j
    End With
    Next i
End If
End Function
'*************************************************************************
'**�� �� ����ClearResource
'**��    �룺(Type_ResourceInSound)Resource
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-08 22:58:12
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub ClearResource(Resource As Type_ResourceInSound)
With Resource
      .ID = -1
      .Unknown = 0
      .strID = ""
End With
End Sub

'*************************************************************************
'**�� �� ����SwapResource
'**��    �룺(Type_ResourceInSound)SoundRes1,(Type_ResourceInSound)SoundRes2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-08 22:15:05
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapResource(Resource1 As Type_ResourceInSound, Resource2 As Type_ResourceInSound)
Dim TemResource As Type_ResourceInSound

TemResource = Resource1
Resource1 = Resource2
Resource2 = TemResource
End Sub

'*************************************************************************
'**�� �� ����SwapInventory
'**��    �룺(Type_XY_Index)Inventory1,(Type_XY_Index)Inventory2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-19 19:53:33
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapInventory(Inventory1 As Type_XY_Index, Inventory2 As Type_XY_Index)
Dim TemInventory As Type_XY_Index

TemInventory = Inventory1
Inventory1 = Inventory2
Inventory2 = TemInventory
End Sub


'*************************************************************************
'**�� �� ����ClearStack
'**��    �룺(Type_Stacks)Stack
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-08 22:58:12
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub ClearStack(Stack As Type_Stacks)
With Stack
      .ID = -1
      .Max = 0
      .Min = 0
      .Flags = 0
End With
End Sub

'*************************************************************************
'**�� �� ����SwapStacks
'**��    �룺(Type_Stacks)Stack1,(Type_Stacks)Stack2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-08 22:15:05
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapStacks(Stack1 As Type_Stacks, Stack2 As Type_Stacks)
Dim TemStack As Type_Stacks

TemStack = Stack1
Stack1 = Stack2
Stack2 = TemStack
End Sub

'*************************************************************************
'**�� �� ����SwapMapIcons
'**��    �룺(Long)MapIcon1,(Long)MapIcon2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-11 00:46:46
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapMapIcons(ByVal MapIcon1 As Long, ByVal MapIcon2 As Long)
Dim TemMapIcon As Type_MapIcon

TemMapIcon = MapIcons(MapIcon1)
 MapIcons(MapIcon1) = MapIcons(MapIcon2)
 MapIcons(MapIcon2) = TemMapIcon
 
SwapLong MapIcons(MapIcon1).ID, MapIcons(MapIcon2).ID
End Sub

'*************************************************************************
'**�� �� ����SwapSounds
'**��    �룺(Long)Sound1,(Long)Sound2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-03 14:21:40
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapSounds(ByVal Sound1 As Long, ByVal Sound2 As Long)
Dim TemSound As Type_Sound

TemSound = Sounds(Sound1)
 Sounds(Sound1) = Sounds(Sound2)
 Sounds(Sound2) = TemSound
 
SwapLong Sounds(Sound1).ID, Sounds(Sound2).ID
End Sub

'*************************************************************************
'**�� �� ����SwapSoundRess
'**��    �룺(Long)SoundRes1,(Long)SoundRes2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-03 14:21:40
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapSoundRess(ByVal SoundRes1 As Long, ByVal SoundRes2 As Long)
Dim TemSoundRes As Type_SoundResource

TemSoundRes = SoundRess(SoundRes1)
 SoundRess(SoundRes1) = SoundRess(SoundRes2)
 SoundRess(SoundRes2) = TemSoundRes
 
SwapLong SoundRess(SoundRes1).ID, SoundRess(SoundRes2).ID
End Sub

'*************************************************************************
'**�� �� ����SwapLong
'**��    �룺(Long)Long1,(Long)Long2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-14 23:20:21
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapLong(Long1 As Long, Long2 As Long)
Dim TemLong As Long

TemLong = Long1
 Long1 = Long2
 Long2 = TemLong
End Sub

'*************************************************************************
'**�� �� ����SwapString
'**��    �룺(String)String1,(String)String2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-06-08 10:57:43
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapString(String1 As String, String2 As String)
Dim TemString As String

TemString = String1
 String1 = String2
 String2 = TemString
End Sub


'*************************************************************************
'**�� �� ����SwapSingle
'**��    �룺(Single)Single1,(Single)Single2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-06-08 10:58:22
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapSingle(Single1 As Single, Single2 As Single)
Dim TemSingle As Single

TemSingle = Single1
Single1 = Single2
Single2 = TemSingle
End Sub

'*************************************************************************
'**�� �� ����SwapItems
'**��    �룺(Long)Item1,(Long)Item2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-14 23:41:42
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapItems(ByVal Item1 As Long, ByVal Item2 As Long)
Dim TemItem As Type_Item

TemItem = itm(Item1)
 itm(Item1) = itm(Item2)
 itm(Item2) = TemItem
 
 SwapLong itm(Item1).ID, itm(Item2).ID
End Sub

'*************************************************************************
'**�� �� ����SwapPSys
'**��    �룺(Long)PSys1,(Long)PSys2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-1-7 14:44:11
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapPSys(ByVal PSys1 As Long, ByVal PSys2 As Long)
Dim TemPSys As Type_Particle_System

TemPSys = PSys(PSys1)
 PSys(PSys1) = PSys(PSys2)
 PSys(PSys2) = TemPSys
 
 SwapLong PSys(PSys1).ID, PSys(PSys2).ID
End Sub

'*************************************************************************
'**�� �� ����SwapMesh
'**��    �룺(Long)Mesh1,(Long)Mesh2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-1-29 21:09:10
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapMesh(ByVal Mesh1 As Long, ByVal Mesh2 As Long)
Dim TemMesh As Type_Mesh

TemMesh = Mesh(Mesh1)
 Mesh(Mesh1) = Mesh(Mesh2)
 Mesh(Mesh2) = TemMesh
 
 SwapLong Mesh(Mesh1).ID, Mesh(Mesh2).ID
End Sub
'*************************************************************************
'**�� �� ����SwapTabMat
'**��    �룺(Long)TabMat1,(Long)TabMat2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-1-23 18:01:32
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapTabMat(ByVal tabmat1 As Long, ByVal tabmat2 As Long)
Dim Temtabmat As Type_Tableau_Material

Temtabmat = TabMat(tabmat1)
 TabMat(tabmat1) = TabMat(tabmat2)
 TabMat(tabmat2) = Temtabmat
 
 SwapLong TabMat(tabmat1).ID, TabMat(tabmat2).ID
End Sub
'*************************************************************************
'**�� �� ����SwapTimeTrg
'**��    �룺(Long)TimeTrg1,(Long)TimeTrg2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-3-4 23:01:12
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapTimeTrg(ByVal TimeTrg1 As Long, ByVal TimeTrg2 As Long)
Dim TemTimeTrg As Type_Time_Trigger

TemTimeTrg = TimeTrg(TimeTrg1)
 TimeTrg(TimeTrg1) = TimeTrg(TimeTrg2)
 TimeTrg(TimeTrg2) = TemTimeTrg
 
 SwapLong TimeTrg(TimeTrg1).ID, TimeTrg(TimeTrg2).ID
End Sub
'*************************************************************************
'**�� �� ����SwapScenes
'**��    �룺(Long)Scene1,(Long)Scene2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-09 23:53:34
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapScenes(ByVal Scene1 As Long, ByVal Scene2 As Long)
Dim TemScene As Type_Scene

TemScene = Scenes(Scene1)
 Scenes(Scene1) = Scenes(Scene2)
 Scenes(Scene2) = TemScene
 
 SwapLong Scenes(Scene1).ID, Scenes(Scene2).ID
End Sub

'*************************************************************************
'**�� �� ����SwapParties
'**��    �룺(Long)Party1,(Long)Party2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-08 23:22:29
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapParties(ByVal Party1 As Long, ByVal Party2 As Long)
Dim TemParty As Type_Party

TemParty = Parties(Party1)
 Parties(Party1) = Parties(Party2)
 Parties(Party2) = TemParty
 
 SwapLong Parties(Party1).ID, Parties(Party2).ID
End Sub

'*************************************************************************
'**�� �� ����SwapPTs
'**��    �룺(Long)PT1,(Long)PT2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-07 16:58:44
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapPTs(ByVal PT1 As Long, ByVal PT2 As Long)
Dim TemPT As Type_PT

TemPT = PTs(PT1)
 PTs(PT1) = PTs(PT2)
 PTs(PT2) = TemPT
 
 SwapLong PTs(PT1).ID, PTs(PT2).ID
End Sub

'*************************************************************************
'**�� �� ����SwapTroops
'**��    �룺(Long)Troop1,(Long)Troop2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-29 22:31:27
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapTroops(ByVal Troop1 As Long, ByVal Troop2 As Long)
Dim TemTroop As Type_Troops

TemTroop = Trps(Troop1)
 Trps(Troop1) = Trps(Troop2)
 Trps(Troop2) = TemTroop
 
  SwapLong Trps(Troop1).ID, Trps(Troop2).ID
End Sub

'*************************************************************************
'**�� �� ����SwapFactions
'**��    �룺(Long)Faction1,(Long)Faction2
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-04 21:37:01
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapFactions(ByVal Faction1 As Long, ByVal Faction2 As Long)
Dim TemFaction As Type_Faction

TemFaction = Factions(Faction1)
 Factions(Faction1) = Factions(Faction2)
 Factions(Faction2) = TemFaction
 
SwapLong Factions(Faction1).ID, Factions(Faction2).ID
End Sub

'*************************************************************************
'**�� �� ����RedimFactions
'**��    �룺(Long)NewCount
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-04 21:52:31
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub RedimFactions(ByVal NewCount As Long)
Dim i As Long

N_Faction = NewCount
ReDim Preserve Factions(N_Faction - 1)

End Sub


'*************************************************************************
'**�� �� ����SwapListItem
'**��    �룺(ListItem)Item1,(ListItem)Item2,(Integer)SubItemCount
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-29 22:33:52
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SwapListItem(ByVal Item1 As ListItem, ByVal Item2 As ListItem, Optional SubItemCount = 0, Optional NoTextSwap As Boolean = False)
Dim strTem As String, i As Integer

If Not NoTextSwap Then
strTem = Item1.Text
Item1.Text = Item2.Text
Item2.Text = strTem
End If

For i = 1 To SubItemCount    'item1.ListSubItems.Count
     strTem = Item1.SubItems(i)
     Item1.SubItems(i) = Item2.SubItems(i)
     Item2.SubItems(i) = strTem

Next i
End Sub

'*************************************************************************
'**�� �� ����OutputItemLine
'**��    �룺(Long)Itm_Idx[��Ϊ-1��ָ��ǰItem]
'**��    ����(String)
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�
'**��    �ڣ�
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-10-49 10:49:26
'**��    ����V1.1321
'*************************************************************************
Public Function OutputItemLine(ByVal Itm_Idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim ItemVes As Type_Item, temLine As String, tmpStr As String, m As Integer

If Itm_Idx >= 0 Then
    ItemVes = itm(Itm_Idx)
Else
    ItemVes = CurrentItm
End If

        With ItemVes
        
            'Check Value
            If Left$(.dbName, 4) <> "itm_" Then
                .dbName = "itm_" & .dbName
            End If
            
            .texname = .disname
            temLine = " " & .dbName & " " & .disname & " " & .texname & " " & CStr(.nmdl) & " "

            For m = 1 To .nmdl
                temLine = temLine & " " & .mdlname(m) & " " & .mdl_b(m) & " "
            Next m
            temLine = temLine & " " & .itmType _
                      & " " & .Action _
                      & " " & CStr(.price) _
                      & " " & CStr(.Prefix) _
                      & " " & Format$(.weight, "0.000000") _
                      & " " & CStr(.abundance) _
                      & " " & CStr(.head_armor) _
                      & " " & CStr(.body_armor) _
                      & " " & CStr(.leg_armor) _
                      & " " & CStr(.difficulty) _
                      & " " & CStr(.hit_points) _
                      & " " & CStr(.speed_rating) _
                      & " " & CStr(.missile_speed) _
                      & " " & CStr(.weapon_length) _
                      & " " & CStr(.max_ammo) _
                      & " " & CStr(.thrust_damage) _
                      & " " & CStr(.swing_damage) & vbCrLf

            temLine = temLine & " " & CStr(.FactionCount) & vbCrLf

            For m = 1 To .FactionCount
                temLine = temLine & " " & .Faction(m).ID
            Next m
            
            If .FactionCount > 0 Then
                temLine = temLine & vbCrLf
            End If
            
            temLine = temLine & .TriggerCount
            
            If .TriggerCount > 0 Then
            temLine = temLine & vbCrLf
               For m = 1 To .TriggerCount
                 tmpStr = OutputTriggerLine(.Trigger(m))
                 temLine = temLine & tmpStr
               Next m
            End If

        End With

     OutputItemLine = temLine & vbCrLf

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputItemLine", Err.Number, Err.Description)
End Function
'*************************************************************************
'**�� �� ����OutputTimeTriggerLine
'**��    �룺(Long)TimeTrg_Idx[��Ϊ-1��ָ��ǰTimeTrg]
'**��    ����(String)
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-02-23 14:09:19
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function OutputTimeTriggerLine(ByVal TimeTrg_idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim TimeTrgVes As Type_Time_Trigger, temLine As String, tmpStr As String, m As Integer

If TimeTrg_idx >= 0 Then
    TimeTrgVes = TimeTrg(TimeTrg_idx)
Else
    TimeTrgVes = CurrentTimeTrg
End If

        With TimeTrgVes

            temLine = CStr(Format(.Check_Interval, "0.000000")) & " " & CStr(Format(.Delay_Interval, "0.000000")) & " " & CStr(Format(.Rearm_Interval, "0.000000")) & "  "
            
           '----------------------������--------------------------
            temLine = temLine & .ConditionsCount & " "
            If .ConditionsCount > 0 Then
               For m = 1 To .ConditionsCount
                 tmpStr = OutputOperationLine(.Condition(m))
                 temLine = temLine & tmpStr
               Next m
            End If
            temLine = temLine & " "
           '------------------------------------------------------
            
           '----------------------�����--------------------------
            temLine = temLine & .ConsequencesCount & " "
            If .ConsequencesCount > 0 Then
               For m = 1 To .ConsequencesCount
                 tmpStr = OutputOperationLine(.Consequence(m))
                 temLine = temLine & tmpStr
               Next m
            End If
           '------------------------------------------------------
           
        End With

     OutputTimeTriggerLine = temLine & vbCrLf

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputTimeTriggerLine", Err.Number, Err.Description)
End Function
'*************************************************************************
'**�� �� ����OutputPSysLine
'**��    �룺(Long)PSys_Idx[��Ϊ-1��ָ��ǰPSys]
'**��    ����(String)
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�
'**��    �ڣ�
'**�� �� �ˣ�Ser_Charles
'**��    �ڣ�2011-1-7 15:53:43
'**��    ����V1.1321
'*************************************************************************
Public Function OutputPSysLine(ByVal PSys_Idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim PSysVes As Type_Particle_System, temLine As String, tmpStr As String, m As Integer

If PSys_Idx >= 0 Then
    PSysVes = PSys(PSys_Idx)
Else
    PSysVes = CurrentPSys
End If

        With PSysVes
        
            'Check Value
            If Left$(.strID, 5) <> "psys_" Then
                .strID = "psys_" & .strID
            End If
            
            temLine = .strID & " " & .Flags & " " & .Mesh_Name & "  " & .Particles_Num & " " & Format(.Life, "0.000000") & " " & Format(.Damping, "0.000000") & " " & Format(.Gravity, "0.000000") & " " & Format(.Turbulance_SZ, "0.000000") & " " & Format(.Turbulance_Str, "0.000000") & " " & vbCrLf

            temLine = temLine & Format(.Alphak(0).X, "0.000000") & " " & Format(.Alphak(0).Y, "0.000000") & "   " & Format(.Alphak(1).X, "0.000000") & " " & Format(.Alphak(1).Y, "0.000000") & vbCrLf _
                      & Format(.Redk(0).X, "0.000000") & " " & Format(.Redk(0).Y, "0.000000") & "   " & Format(.Redk(1).X, "0.000000") & " " & Format(.Redk(1).Y, "0.000000") & vbCrLf _
                      & Format(.Greenk(0).X, "0.000000") & " " & Format(.Greenk(0).Y, "0.000000") & "   " & Format(.Greenk(1).X, "0.000000") & " " & Format(.Greenk(1).Y, "0.000000") & vbCrLf _
                      & Format(.Bluek(0).X, "0.000000") & " " & Format(.Bluek(0).Y, "0.000000") & "   " & Format(.Bluek(1).X, "0.000000") & " " & Format(.Bluek(1).Y, "0.000000") & vbCrLf _
                      & Format(.Scalek(0).X, "0.000000") & " " & Format(.Scalek(0).Y, "0.000000") & "   " & Format(.Scalek(1).X, "0.000000") & " " & Format(.Scalek(1).Y, "0.000000") & vbCrLf _
                      & Format(.EBSZ(0), "0.000000") & " " & Format(.EBSZ(1), "0.000000") & " " & Format(.EBSZ(2), "0.000000") & "   " _
                      & Format(.EV(0), "0.000000") & " " & Format(.EV(1), "0.000000") & " " & Format(.EV(2), "0.000000") & "   " _
                      & Format(.EDR, "0.000000") & " " & vbCrLf _
                      & Format(.PRS, "0.000000") & " " & Format(.PRD, "0.000000") & " "
         End With
         
     OutputPSysLine = temLine

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputPSysLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����OutputTabMatLine
'**��    �룺(Long)TabMat_Idx[��Ϊ-1��ָ��ǰ�ɱ��ز�]
'**��    ����(String)
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�
'**��    �ڣ�
'**�� �� �ˣ�Ser_Charles
'**��    �ڣ�2011-1-22 19:18:22
'**��    ����V1.1321
'*************************************************************************
Public Function OutputTabMatLine(ByVal TabMat_Idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim TabMatVes As Type_Tableau_Material, temLine As String, tmpStr As String, m As Integer, i As Integer

If TabMat_Idx >= 0 Then
    TabMatVes = TabMat(TabMat_Idx)
Else
    TabMatVes = CurrentTabMat
End If

        With TabMatVes
        
            'Check Value
            If Left$(.strID, 4) <> "tab_" Then
                .strID = "tab_" & .strID
            End If
            
            temLine = .strID & " " & .Flags & " " & .Sample & " " & .Width & " " & .Height & " " & .Min.X & " " & .Min.Y & " " & .Max.X & " " & .Max.Y & " "

            temLine = temLine & .OpCount & " "
            
            If .OpCount > 0 Then
               For i = 1 To .OpCount
                 'tmpStr = OutputTriggerLine(.trigger(m))
                 temLine = temLine & .OpBlock(i).Op & " " & .OpBlock(i).ParaNum & " "
                 If .OpBlock(i).ParaNum > 0 Then
                       For m = 1 To .OpBlock(i).ParaNum
                            With .OpBlock(i)
                                temLine = temLine & .Para(m).Value & " "
                            End With
                       Next m
                 End If
               Next i
            End If
            
         End With
         
     OutputTabMatLine = temLine

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputTabMatLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����OutputMeshLine
'**��    �룺(Long)Mesh_Idx[��Ϊ-1��ָ��ǰMesh]
'**��    ����(String)
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�
'**��    �ڣ�
'**�� �� �ˣ�Ser_Charles
'**��    �ڣ�2011-1-28 19:19:55
'**��    ����V1.1321
'*************************************************************************
Public Function OutputMeshLine(ByVal Mesh_Idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim MeshVes As Type_Mesh, temLine As String, tmpStr As String, m As Integer

If Mesh_Idx >= 0 Then
    MeshVes = Mesh(Mesh_Idx)
Else
    MeshVes = CurrentMesh
End If

        With MeshVes
        
            'Check Value
            If Left$(.strID, 5) <> "mesh_" Then
                .strID = "mesh_" & .strID
            End If
            
            temLine = .strID & " " & .Flags & " " & .Resource_Name & " "

            temLine = temLine & Format(.Translation.X, "0.000000") & " " & Format(.Translation.Y, "0.000000") & " " & Format(.Translation.Z, "0.000000") & " " _
                              & Format(.Rotation_Angle.X, "0.000000") & " " & Format(.Rotation_Angle.Y, "0.000000") & " " & Format(.Rotation_Angle.Z, "0.000000") & " " _
                              & Format(.Scale.X, "0.000000") & " " & Format(.Scale.Y, "0.000000") & " " & Format(.Scale.Z, "0.000000")
         End With
         
     OutputMeshLine = temLine

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputMeshLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����OutputTroopLine
'**��    �룺(Long)Trp_Idx[��Ϊ-1��ָ��ǰTroop]
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-30 11:36:44
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function OutputTroopLine(ByVal Trp_Idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim TroopVes As Type_Troops, temLine As String, tmpStr As String, m As Integer

If Trp_Idx >= 0 Then
    TroopVes = Trps(Trp_Idx)
Else
    TroopVes = CurrentTrp
End If

With TroopVes

            'Check Value
            If Left$(.strID, 4) <> "trp_" Then
                .strID = "trp_" & .strID
            End If

            temLine = .strID & " " & .strName & " " & .strPtName & " " & .unknown_warband(1) & " " & .Flags & _
                  " " & CStr(.Scene) & " " & CStr(.reserved) & " " & CStr(.Faction) & " " & .Upgrade1 & " " & .Upgrade2 & vbCrLf
                  
            tmpStr = "  "
            For m = 1 To 64
                tmpStr = tmpStr & CStr(.lstInventory(m).X) & " " & CStr(.lstInventory(m).Y) & " "
            Next m
            temLine = temLine & tmpStr & vbCrLf

            temLine = temLine & "  " & CStr(.tAttrib.strPoint) & " " & CStr(.tAttrib.agiPoint) & _
                  " " & CStr(.tAttrib.intPoint) & " " & CStr(.tAttrib.chaPoint) & " " & CStr(.tAttrib.level) & vbCrLf

            temLine = temLine & " " & .WP.one_handed & " " & .WP.two_handed & " " & .WP.polearm & _
                  " " & .WP.archery & " " & .WP.crossbow & " " & .WP.throwing & " " & .WP.firearm & vbCrLf

            temLine = temLine & "" & CStr(.Skills(1)) & " " & CStr(.Skills(2)) & " " & CStr(.Skills(3)) & _
                  " " & CStr(.Skills(4)) & " " & CStr(.Skills(5)) & " " & CStr(.Skills(6)) & " " & vbCrLf

            temLine = temLine & "  " & .Face(1) & " " & .Face(2) & " " & .Face(3) & " " & .Face(4) & " " & _
                  .Face(5) & " " & .Face(6) & " " & .Face(7) & " " & .Face(8) & " " & vbCrLf

            
        End With
        
        OutputTroopLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputTroopLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����OutputFactionLine
'**��    �룺(Long)Faction_Idx[��Ϊ-1��ָ��ǰFaction]
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-30 11:36:44
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function OutputFactionLine(ByVal Faction_Idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim FactionVes As Type_Faction, temLine As String, tmpStr As String, m As Integer

If Faction_Idx >= 0 Then
    FactionVes = Factions(Faction_Idx)
Else
    FactionVes = CurrentFaction
End If

With FactionVes

            'Check Value
            If Left$(.strID, 4) <> "fac_" Then
                .strID = "fac_" & .strID
            End If

            temLine = .strID & " " & .strName & " " & CStr(.Flags) & " " & CStr(.lColor) & " " & vbCrLf

            tmpStr = " "
            For m = 0 To UBound(.RelationShip)
                tmpStr = tmpStr & Format(.RelationShip(m).Value, "0.000000") & " "
            Next m
            temLine = temLine & tmpStr & vbCrLf

            temLine = temLine & .reserved & " "
            
        End With
        
        OutputFactionLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputFactionLine", Err.Number, Err.Description)
End Function


'*************************************************************************
'**�� �� ����OutputMapIconLine
'**��    �룺(Long)MapIcon_Idx[��Ϊ-1��ָ��ǰMapIcon]
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-11 11:15:06
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function OutputMapIconLine(ByVal MapIcon_Idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim MapIconVes As Type_MapIcon, temLine As String, tmpStr As String, m As Integer

If MapIcon_Idx >= 0 Then
    MapIconVes = MapIcons(MapIcon_Idx)
Else
    MapIconVes = CurrentMapIcon
End If

With MapIconVes


            temLine = .strID & " " & CStr(.Flags) & " " & .MeshName & " " & .mScale & " " & .Sound & " "

            For m = 0 To UBound(.Offset)
               temLine = temLine & .Offset(m) & " "
            Next m

            temLine = temLine & .TriggerCount
            
            If .TriggerCount > 0 Then
            temLine = temLine & vbCrLf
               For m = 1 To .TriggerCount
                 tmpStr = OutputTriggerLine(.Triggers(m))
                 temLine = temLine & tmpStr
               Next m
            End If
            
            temLine = temLine & vbCrLf & vbCrLf
        End With
        
        OutputMapIconLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputMapIconLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����OutputOperationLine
'**��    �룺(Type_Op_Block)OperationVes
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-11 11:13:10
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function OutputOperationLine(OperationVes As Type_Op_Block) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim temLine As String, m As Integer

With OperationVes

            temLine = .Op & " " & .ParaNum & " "
         
       If .ParaNum > 0 Then
            For m = 1 To .ParaNum
                temLine = temLine & .Para(m).Value & " "
            Next m
       End If
            
End With
        
        OutputOperationLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputOperationLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����OutputTriggerLine
'**��    �룺(Type_Op_Block)TriggerVes
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-11 11:21:05
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function OutputTriggerLine(TriggerVes As Type_Trigger) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim temLine As String, tmpStr As String, m As Integer

With TriggerVes

            temLine = Format(.tiOn, "0.000000") & "  " & .ActNum & " "
         
       If .ActNum > 0 Then
            For m = 1 To .ActNum
                tmpStr = OutputOperationLine(.tiAct(m))
                temLine = temLine & tmpStr
            Next m
       End If
            
End With
        
        OutputTriggerLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputTriggerLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����OutputPartyLine
'**��    �룺(Long)Party_Idx[��Ϊ-1��ָ��ǰParty]
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-30 11:36:44
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function OutputPartyLine(ByVal Party_Idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim PartyVes As Type_Party, temLine As String, tmpStr As String, m As Integer

If Party_Idx >= 0 Then
    PartyVes = Parties(Party_Idx)
Else
    PartyVes = CurrentParty
End If

With PartyVes
            'Check Value
            If Left$(.strID, 2) <> "p_" Then
                .strID = "p_" & .strID
            End If
                temLine = " " & (.UnknownTitle) & " " & (.ID) & " " & (.id2) & " " & (.strID) & " " & (.strName) & " " & (.Flags) & " " & (.Menu) & " " & (.Template) & " " & (.Faction) & " "
                
                For m = 1 To 2
                     temLine = temLine & .Personality(m) & " "
                Next m
                
                temLine = temLine & .AI_Behavior & " " & .AI_Target & " " & .reserved & " "
                
                For m = 1 To 3
                     temLine = temLine & .InitPos(m).X & " " & .InitPos(m).Y & " "
                Next m
                
                temLine = temLine & .UnknownStr & " " & .StacksCount & " "
                
                
                If .StacksCount > 0 Then
                     For m = 1 To .StacksCount
                          temLine = temLine & .Stacks(m).ID & " "
                          temLine = temLine & .Stacks(m).Min & " "
                          temLine = temLine & .Stacks(m).Max & " "
                          temLine = temLine & .Stacks(m).Flags & " "
                     Next m
                End If
                
                temLine = temLine & vbCrLf & .Degree
                
            End With
        OutputPartyLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputPartyLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����OutputSceneLine
'**��    �룺(Long)Scene_Idx[��Ϊ-1��ָ��ǰScene]
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-10 22:20:38
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function OutputSceneLine(ByVal Scene_Idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim SceneVes As Type_Scene, temLine As String, tmpStr As String, m As Integer

If Scene_Idx >= 0 Then
    SceneVes = Scenes(Scene_Idx)
Else
    SceneVes = CurrentScene
End If

With SceneVes
            'Check Value
            If Left$(.strID, 4) <> "scn_" Then
                .strID = "scn_" & .strID
            End If
                temLine = (.strID) & " " & (.strName) & " " & (.Flags) & " " & (.MeshName) & " " & (.BodyName) & " "
                
                For m = 0 To 1
                     temLine = temLine & Format(.p(m).X, "0.000000") & " " & Format(.p(m).Y, "0.000000") & " "
                Next m
                
                temLine = temLine & Format(.WaterLevel, "0.000000") & " " & .TerrainCode & " " & vbCrLf
                
                temLine = temLine & "  " & .AccessCount & " "
                
                If .AccessCount > 0 Then
                     For m = 1 To .AccessCount
                          temLine = temLine & " " & .Accesses(m) & " "
                     Next m
                End If
                
                temLine = temLine & vbCrLf & "  " & .ChestCount & " "
                
                If .ChestCount > 0 Then
                     For m = 1 To .ChestCount
                          temLine = temLine & " " & .Chests(m).ID & " "
                     Next m
                End If
                temLine = temLine & vbCrLf & " " & .Outer_Terrain_Type & " "
                
            End With
        OutputSceneLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputSceneLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����OutputSoundLine
'**��    �룺(Long)Sound_Idx[��Ϊ-1��ָ��ǰSound]
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-10 22:20:38
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function OutputSoundLine(ByVal Sound_Idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim SoundVes As Type_Sound, temLine As String, tmpStr As String, m As Integer

If Sound_Idx >= 0 Then
    SoundVes = Sounds(Sound_Idx)
Else
    SoundVes = CurrentSound
End If

With SoundVes
            'Check Value
            If Left$(.sndName, 4) <> "snd_" Then
                .sndName = "snd_" & .sndName
            End If
                temLine = (.sndName) & " " & (.Flags)
                
                temLine = temLine & " " & .ResourceCount & " "
                
                If .ResourceCount > 0 Then
                     For m = 1 To .ResourceCount
                          temLine = temLine & .Resource(m).ID & " " & .Resource(m).Unknown & " "
                     Next m
                End If
                
            End With
        OutputSoundLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputSoundLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����OutputSoundResLine
'**��    �룺(Long)SoundRes_Idx[��Ϊ-1��ָ��ǰSoundRes]
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-03 15:52:59
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function OutputSoundResLine(ByVal SoundRes_Idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim SoundResVes As Type_SoundResource, temLine As String, m As Integer

If SoundRes_Idx >= 0 Then
    SoundResVes = SoundRess(SoundRes_Idx)
Else
    SoundResVes = CurrentSoundRes
End If

With SoundResVes

                temLine = " " & (.sndName) & " " & (.Flags)
                
            End With
        OutputSoundResLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputSoundResLine", Err.Number, Err.Description)
End Function


'*************************************************************************
'**�� �� ����OutputPTLine
'**��    �룺(Long)PT_Idx[��Ϊ-1��ָ��ǰPT]
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-30 11:36:44
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function OutputPTLine(ByVal PT_Idx As Long) As String
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim PTVes As Type_PT, temLine As String, tmpStr As String, m As Integer

If PT_Idx >= 0 Then
    PTVes = PTs(PT_Idx)
Else
    PTVes = CurPartyTemplate
End If

With PTVes
            'Check Value
            If Left$(.ptID, 3) <> "pt_" Then
                .ptID = "pt_" & .ptID
            End If
                temLine = (.ptID) & " " & (.ptName) & " " & (.Flags) & " " & (.Menu) & " " & (.Faction) & " " & (.Personality) & " "

                For m = 1 To 6
                    If .Stacks(m).ID < 0 Then
                        temLine = temLine & ("-1 ")
                    Else
                        temLine = temLine & (.Stacks(m).ID) & " "
                      If .Stacks(m).ID >= 0 Then
                        temLine = temLine & (.Stacks(m).Min) & " "

                        temLine = temLine & (.Stacks(m).Max) & " "
                        
                        temLine = temLine & (.Stacks(m).Flags) & " "
                      End If
                    End If
                Next
                
            End With
        OutputPTLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "OutputPTLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����CheckEditable
'**��    �룺(Long)QueryType,(Long)Index
'**��    ����(Boolean)
'**����������
'**ȫ�ֱ�����N_Item
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-30 14:16:59
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function CheckEditable(ByVal QueryType As Long, ByVal Index As Long) As Boolean

If QueryType <> EditInfo_ItemsCount Then
    CheckEditable = Index >= MnBInfo.EditInfo(QueryType)
    '������2013-06-28, ȥ��ɾ������
    CheckEditable = True
Else
    CheckEditable = (Index >= MnBInfo.EditInfo(QueryType) - 1) And (Index < N_Item - 1)
    '������2013-06-28, ȥ��ɾ������
    CheckEditable = Index < N_Item - 1
End If

End Function

'*************************************************************************
'**�� �� ����CheckExist
'**��    �룺(Long)QueryType,(Long)Index
'**��    ����(Boolean)
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-30 16:30:09
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function CheckExist(ByVal QueryType As Long, ByVal Index As Long) As Boolean

Select Case QueryType     'A_E
       Case EditInfo_TroopsCount
            CheckExist = Index > -1 And Index < N_Troop
       Case EditInfo_FactionsCount
            CheckExist = Index > -1 And Index < N_Faction
       Case EditInfo_ItemsCount
            CheckExist = Index > -1 And Index < N_Item
       Case EditInfo_PartiesCount
            CheckExist = Index > -1 And Index < N_Party
       Case EditInfo_PartyTemplatesCount
            CheckExist = Index > -1 And Index < N_PT
       Case EditInfo_ScenesCount
            CheckExist = Index > -1 And Index < N_Scene
       Case EditInfo_MapIconsCount
            CheckExist = Index > -1 And Index < N_MapIcon
       Case EditInfo_SoundsCount
            CheckExist = Index > -1 And Index < N_Sound
       Case EditInfo_SoundRessCount
            CheckExist = Index > -1 And Index < N_SoundRes
       Case EditInfo_PSysCount
            CheckExist = Index > -1 And Index < N_PSys
       Case EditInfo_TabMatCount
            CheckExist = Index > -1 And Index < N_TabMat
       Case EditInfo_MeshCount
            CheckExist = Index > -1 And Index < N_Mesh
       Case EditInfo_TimeTrgCount
            CheckExist = Index > -1 And Index < N_TimeTrg
       Case EditInfo_StringsCount
            CheckExist = Index > -1 And Index < N_Str
End Select
End Function

'*************************************************************************
'**�� �� ����SaveItemCSVFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-30 22:08:41
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SaveItemCSVFile(FilePath As String)
     On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim strTem As String, i As Long, q As Boolean
    
    If LCase$(MnBInfo.Language) = "en" Then
         'SaveItemCSVFile = True
         Exit Sub
    End If
    
    For i = 0 To N_Item - 1
       If itm(i).csvName <> itm(i).disname Then
          strTem = strTem & itm(i).dbName & "|" & itm(i).csvName & vbCrLf
       End If
       
       If itm(i).csvName_pl <> itm(i).disname Then
          strTem = strTem & itm(i).dbName & "_pl" & "|" & itm(i).csvName_pl & vbCrLf
       End If
    Next i
    
    q = UEFSaveTextFile(FilePath, strTem, False, UEF_UTF8, UEF_UTF8)
    
    If q = False Then
        Call logErr("ModMain", "SaveItemCSVFile", "INVALID_HANDLE_VALUE", "�޷����ļ�:[" & FilePath & "]")
    End If
    
     Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveItemCSVFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SaveStringCSVFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-30 22:08:41
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SaveStringCSVFile(FilePath As String)
     On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim strTem As String, i As Long, q As Boolean
    
    If LCase$(MnBInfo.Language) = "en" Then
         'SaveStringCSVFile = True
         Exit Sub
    End If
    
    For i = 0 To N_Str - 1
       If Strs(i).CSV <> "" Then
          strTem = strTem & Strs(i).Name & "|" & Strs(i).CSV
          If i <> N_Str - 1 Then
             strTem = strTem & vbCrLf
          End If
       End If
    Next i
    
    q = UEFSaveTextFile(FilePath, strTem, False, UEF_UTF8, UEF_UTF8)
    
    If q = False Then
        Call logErr("ModMain", "SaveStringCSVFile", "INVALID_HANDLE_VALUE", "�޷����ļ�:[" & FilePath & "]")
    End If
    
     Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveStringCSVFile", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����SaveFactionCSVFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-30 23:24:31
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SaveFactionCSVFile(FilePath As String)
     On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim strTem As String, i As Long, q As Boolean
    
    If LCase$(MnBInfo.Language) = "en" Then
         'SaveFactionCSVFile = True
         Exit Sub
    End If
    
    For i = 0 To N_Faction - 1
       If Factions(i).csvName <> Factions(i).strName Then
          strTem = strTem & Factions(i).strID & "|" & Factions(i).csvName & vbCrLf
       End If
       
    Next i
    
    q = UEFSaveTextFile(FilePath, strTem, False, UEF_UTF8, UEF_UTF8)
    
    If q = False Then
        Call logErr("ModMain", "SaveFactionCSVFile", "INVALID_HANDLE_VALUE", "�޷����ļ�:[" & FilePath & "]")
    End If
    
     Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveFactionCSVFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SaveFactionFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�
'**��    �ڣ�
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-12-04 23:34:05
'**��    ����V1.1321
'*************************************************************************
Sub SaveFactionFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim tmpStr As String
    Dim n As Long, i As Long
    Dim m As Integer

    lngHandle = FreeFile()

    Open FilePath For Output As #lngHandle

    tmpStr = ""
    For n = 0 To 1
        tmpStr = tmpStr & FactionVersionInform(n) & " "
    Next n
    tmpStr = tmpStr & FactionVersionInform(2)
    Print #lngHandle, tmpStr
    
    Print #lngHandle, Trim$(Str$(N_Faction))
    
    tmpStr = ""
    For n = 0 To N_Faction - 1
            tmpStr = tmpStr & OutputFactionLine(n)
    Next n
    Print #lngHandle, tmpStr
    Close #lngHandle

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveFactionFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����SaveMapIconFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�
'**��    �ڣ�
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-12-11 11:28:20
'**��    ����V1.1321
'*************************************************************************
Sub SaveMapIconFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim tmpStr As String
    Dim n As Long, i As Long
    Dim m As Integer

    lngHandle = FreeFile()

    Open FilePath For Output As #lngHandle

    tmpStr = ""
    For n = 0 To 1
        tmpStr = tmpStr & MapIconVersionInform(n) & " "
    Next n
    tmpStr = tmpStr & MapIconVersionInform(2)
    Print #lngHandle, tmpStr
    
    Print #lngHandle, Trim$(Str$(N_MapIcon))
    
    For n = 0 To N_MapIcon - 1
            tmpStr = OutputMapIconLine(n)
            Print #lngHandle, tmpStr
    Next n
   
    Close #lngHandle

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "SaveMapIconFile", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**�� �� ����LoadIModFile
'**��    �룺FileName(String) -
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-01 23:14:15
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Function LoadIModFile(FileName As String) As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    LoadIModFile = False
    
    Dim tmpFileName As String
    tmpFileName = FileName
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Function
    End If
    
    Dim TemP As String, arrTmp() As String, n As Long, m As Integer

    n = 0
    
    Dim arrFileBuff() As String
    Dim i As Long
    m = FreeFile
    ReDim arrTmp(0)
    Open tmpFileName For Input As #m
         Do While Not (EOF(m))
            Input #m, arrTmp(n)
                    For i = 1 To Len(arrTmp(n))
                        If Mid(arrTmp(n), i, 1) = " " Then
                            arrTmp(n) = Left(arrTmp(n), i - 1)
                            Exit For
                        End If
                    Next i
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                n = n + 1
         Loop
    Close #m
    
    Do While arrTmp(UBound(arrTmp)) = ""
        ReDim Preserve arrTmp(UBound(arrTmp) - 1)
    Loop
    
    N_IMod = UBound(arrTmp) + 1
    
    ReDim IMod(N_IMod - 1)
    
    For i = 0 To N_IMod - 1
        IMod(i).ID = arrTmp(i)
    Next i
    
    LoadIModFile = True

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadIModFile", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����LoadPSysFile
'**��    �룺filePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-09 14:32:22
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub LoadPSysFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    For n = 0 To 2
        PSysVersionInform(n) = GetWord()
    Next n
    N_PSys = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_PSysCount, CStr(N_PSys)
        MnBInfo.EditInfo(EditInfo_PSysCount) = N_PSys
    End If
    
    ReDim PSys(N_PSys - 1)

    DoEvents
    For n = 0 To N_PSys - 1

        With PSys(n)
            .ID = n
            .strID = GetWord()
            .Flags = GetWord()
            .Mesh_Name = GetWord()
            .Particles_Num = CLng(Val(GetWord()))
            .Life = Val(GetWord())
            .Damping = Val(GetWord())
            .Gravity = Val(GetWord())
            .Turbulance_SZ = Val(GetWord())
            .Turbulance_Str = Val(GetWord())
            
            .Alphak(0).X = Val(GetWord())
            .Alphak(0).Y = Val(GetWord())
            .Alphak(1).X = Val(GetWord())
            .Alphak(1).Y = Val(GetWord())
            .Redk(0).X = Val(GetWord())
            .Redk(0).Y = Val(GetWord())
            .Redk(1).X = Val(GetWord())
            .Redk(1).Y = Val(GetWord())
            .Greenk(0).X = Val(GetWord())
            .Greenk(0).Y = Val(GetWord())
            .Greenk(1).X = Val(GetWord())
            .Greenk(1).Y = Val(GetWord())
            .Bluek(0).X = Val(GetWord())
            .Bluek(0).Y = Val(GetWord())
            .Bluek(1).X = Val(GetWord())
            .Bluek(1).Y = Val(GetWord())
            .Scalek(0).X = Val(GetWord())
            .Scalek(0).Y = Val(GetWord())
            .Scalek(1).X = Val(GetWord())
            .Scalek(1).Y = Val(GetWord())
            
            .EBSZ(0) = Val(GetWord())
            .EBSZ(1) = Val(GetWord())
            .EBSZ(2) = Val(GetWord())
            .EV(0) = Val(GetWord())
            .EV(1) = Val(GetWord())
            .EV(2) = Val(GetWord())
            .EDR = Val(GetWord())
            .PRS = Val(GetWord())
            .PRD = Val(GetWord())
            
            .Edit = CheckEditable(EditInfo_PSysCount, n)
            
            AddIndex .ID, .strID
        End With
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadPSysFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����LoadTabMatFile
'**��    �룺filePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-09 14:32:22
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub LoadTabMatFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim i As Integer
    Dim m As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    'For n = 0 To 2
    '    TabMatVersionInform(n) = GetWord()
    'Next n
    N_TabMat = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_TabMatCount, CStr(N_TabMat)
        MnBInfo.EditInfo(EditInfo_TabMatCount) = N_TabMat
    End If
    
    ReDim TabMat(N_TabMat - 1)

    DoEvents
    For n = 0 To N_TabMat - 1

        With TabMat(n)
            .ID = n
            .strID = GetWord()
            .Flags = GetWord()
            .Sample = GetWord()
            .Width = CLng(Val(GetWord()))
            .Height = CLng(Val(GetWord()))
            .Min.X = CLng(Val(GetWord()))
            .Min.Y = CLng(Val(GetWord()))
            .Max.X = CLng(Val(GetWord()))
            .Max.Y = CLng(Val(GetWord()))
            .OpCount = CLng(Val(GetWord()))
            
            If .OpCount = 0 Then
                 ReDim .OpBlock(0)
            Else
                 ReDim .OpBlock(1 To .OpCount)
                 For i = 1 To .OpCount
                       .OpBlock(i).Op = GetWord()
                       .OpBlock(i).ParaNum = CLng(Val(GetWord))
                       If .OpBlock(i).ParaNum = 0 Then
                            ReDim .OpBlock(i).Para(0)
                       Else
                            ReDim .OpBlock(i).Para(1 To .OpBlock(i).ParaNum)
                            For m = 1 To .OpBlock(i).ParaNum
                                 .OpBlock(i).Para(m).Value = GetWord()
                            Next m
                       End If
                 Next i
            End If
            
            .Edit = CheckEditable(EditInfo_TabMatCount, n)
            
            AddIndex .ID, .strID
        End With
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadTabMatFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����LoadMeshFile
'**��    �룺filePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-1-28 19:10:12
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub LoadMeshFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    'For n = 0 To 2
    '    MeshVersionInform(n) = GetWord()
    'Next n
    N_Mesh = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_MeshCount, CStr(N_Mesh)
        MnBInfo.EditInfo(EditInfo_MeshCount) = N_Mesh
    End If
    
    ReDim Mesh(N_Mesh - 1)

    DoEvents
    For n = 0 To N_Mesh - 1

        With Mesh(n)
            .ID = n
            .strID = GetWord()
            .Flags = GetWord()
            .Resource_Name = GetWord()
            
            .Translation.X = Val(GetWord())
            .Translation.Y = Val(GetWord())
            .Translation.Z = Val(GetWord())

            .Rotation_Angle.X = Val(GetWord())
            .Rotation_Angle.Y = Val(GetWord())
            .Rotation_Angle.Z = Val(GetWord())
            
            .Scale.X = Val(GetWord())
            .Scale.Y = Val(GetWord())
            .Scale.Z = Val(GetWord())
            
            .Edit = CheckEditable(EditInfo_MeshCount, n)
            
            AddIndex .ID, .strID
        End With
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadMeshFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����LoadTriggerFile
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-02-22 12:59:11
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub LoadTriggerFile(FilePath As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer, i As Integer, H As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "ȱ���ļ�: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    For n = 0 To 2
        TimeTrgVersionInform(n) = GetWord()
    Next n
    N_TimeTrg = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_TimeTrgCount) = N_TimeTrg
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_TimeTrgCount, CStr(N_TimeTrg)
    End If
    
    ReDim TimeTrg(N_TimeTrg - 1)

    DoEvents
    For n = 0 To N_TimeTrg - 1

        With TimeTrg(n)
            .ID = n
            .Check_Interval = Val(GetWord())
            .Delay_Interval = Val(GetWord())
            .Rearm_Interval = Val(GetWord())
    
    '-----------------------------������------------------------------------------
       .ConditionsCount = CLng(Val(GetWord()))
       If .ConditionsCount > 0 Then
            ReDim .Condition(1 To .ConditionsCount)
            For i = 1 To .ConditionsCount
                       .Condition(i).Op = GetWord()
                       .Condition(i).ParaNum = Val(GetWord())
                       If .Condition(i).ParaNum > 0 Then
                           ReDim .Condition(i).Para(1 To .Condition(i).ParaNum)
                           For H = 1 To .Condition(i).ParaNum
                                 .Condition(i).Para(H).Value = GetWord()
                           Next H
                       Else
                           ReDim .Condition(i).Para(0)
                       End If
            Next i
       ElseIf .ConditionsCount <= 0 Then
            ReDim .Condition(0)
       End If
     '-----------------------------------------------------------------------------
     
    '-----------------------------�����------------------------------------------
       .ConsequencesCount = CLng(Val(GetWord()))
       If .ConsequencesCount > 0 Then
            ReDim .Consequence(1 To .ConsequencesCount)
            For i = 1 To .ConsequencesCount
                       .Consequence(i).Op = GetWord()
                       .Consequence(i).ParaNum = Val(GetWord())
                       If .Consequence(i).ParaNum > 0 Then
                           ReDim .Consequence(i).Para(1 To .Consequence(i).ParaNum)
                           For H = 1 To .Consequence(i).ParaNum
                                 .Consequence(i).Para(H).Value = GetWord()
                           Next H
                       Else
                           ReDim .Consequence(i).Para(0)
                       End If
            Next i
       ElseIf .ConsequencesCount <= 0 Then
            ReDim .Consequence(0)
       End If
     '-----------------------------------------------------------------------------
     
            .Edit = CheckEditable(EditInfo_TimeTrgCount, n)
            
        End With
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "LoadTriggerFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����ReadPSysLine
'**��    �룺Text(String),OutputPSys(Type_Particle_System)
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-09 17:07:02
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub ReadPSysLine(Text As String, OutputPSys As Type_Particle_System)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer
    
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()

        With OutputPSys
        
            .strID = GetWordL()
            .Flags = GetWordL()
            .Mesh_Name = GetWordL()
            .Particles_Num = CLng(Val(GetWordL()))
            .Life = Val(GetWordL())
            .Damping = Val(GetWordL())
            .Gravity = Val(GetWordL())
            .Turbulance_SZ = Val(GetWordL())
            .Turbulance_Str = Val(GetWordL())
            
            .Alphak(0).X = Val(GetWordL())
            .Alphak(0).Y = Val(GetWordL())
            .Alphak(1).X = Val(GetWordL())
            .Alphak(1).Y = Val(GetWordL())
            .Redk(0).X = Val(GetWordL())
            .Redk(0).Y = Val(GetWordL())
            .Redk(1).X = Val(GetWordL())
            .Redk(1).Y = Val(GetWordL())
            .Greenk(0).X = Val(GetWordL())
            .Greenk(0).Y = Val(GetWordL())
            .Greenk(1).X = Val(GetWordL())
            .Greenk(1).Y = Val(GetWordL())
            .Bluek(0).X = Val(GetWordL())
            .Bluek(0).Y = Val(GetWordL())
            .Bluek(1).X = Val(GetWordL())
            .Bluek(1).Y = Val(GetWordL())
            .Scalek(0).X = Val(GetWordL())
            .Scalek(0).Y = Val(GetWordL())
            .Scalek(1).X = Val(GetWordL())
            .Scalek(1).Y = Val(GetWordL())
            
            .EBSZ(0) = Val(GetWordL())
            .EBSZ(1) = Val(GetWordL())
            .EBSZ(2) = Val(GetWordL())
            .EV(0) = Val(GetWordL())
            .EV(1) = Val(GetWordL())
            .EV(2) = Val(GetWordL())
            .EDR = Val(GetWordL())
            .PRS = Val(GetWordL())
            .PRD = Val(GetWordL())
               
        End With
    
    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ReadPSysLine", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����readTriggerLine
'**��    �룺FilePath(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2012-03-13 11:35:36
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub readTriggerLine(Text As String, OutputTrigger As Type_Time_Trigger)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Dim n As Long
    Dim i As Integer
    Dim m As Integer
    Dim H As Integer
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
    
        With OutputTrigger
            .Check_Interval = Val(GetWordL())
            .Delay_Interval = Val(GetWordL())
            .Rearm_Interval = Val(GetWordL())
    
    '-----------------------------������------------------------------------------
       .ConditionsCount = CLng(Val(GetWordL()))
       If .ConditionsCount > 0 Then
            ReDim .Condition(1 To .ConditionsCount)
            For i = 1 To .ConditionsCount
                       .Condition(i).Op = GetWordL()
                       .Condition(i).ParaNum = Val(GetWordL())
                       If .Condition(i).ParaNum > 0 Then
                           ReDim .Condition(i).Para(1 To .Condition(i).ParaNum)
                           For H = 1 To .Condition(i).ParaNum
                                 .Condition(i).Para(H).Value = GetWordL()
                           Next H
                       Else
                           ReDim .Condition(i).Para(0)
                       End If
            Next i
       ElseIf .ConditionsCount <= 0 Then
            ReDim .Condition(0)
       End If
     '-----------------------------------------------------------------------------
     
    '-----------------------------�����------------------------------------------
       .ConsequencesCount = CLng(Val(GetWordL()))
       If .ConsequencesCount > 0 Then
            ReDim .Consequence(1 To .ConsequencesCount)
            For i = 1 To .ConsequencesCount
                       .Consequence(i).Op = GetWordL()
                       .Consequence(i).ParaNum = Val(GetWordL())
                       If .Consequence(i).ParaNum > 0 Then
                           ReDim .Consequence(i).Para(1 To .Consequence(i).ParaNum)
                           For H = 1 To .Consequence(i).ParaNum
                                 .Consequence(i).Para(H).Value = GetWordL()
                           Next H
                       Else
                           ReDim .Consequence(i).Para(0)
                       End If
            Next i
       ElseIf .ConsequencesCount <= 0 Then
            ReDim .Consequence(0)
       End If
     '-----------------------------------------------------------------------------
            
        End With
    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "readTriggerLine [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����ReadTabMatLine
'**��    �룺Text(String),OutputTabMat(Type_Tableau_Material) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-09 14:32:22
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub ReadTabMatLine(Text As String, OutputTabMat As Type_Tableau_Material)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim i As Integer
    Dim m As Integer
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
    
        With OutputTabMat
            .strID = GetWordL()
            .Flags = GetWordL()
            .Sample = GetWordL()
            .Width = CLng(Val(GetWordL()))
            .Height = CLng(Val(GetWordL()))
            .Min.X = CLng(Val(GetWordL()))
            .Min.Y = CLng(Val(GetWordL()))
            .Max.X = CLng(Val(GetWordL()))
            .Max.Y = CLng(Val(GetWordL()))
            .OpCount = CLng(Val(GetWordL()))
            
            If .OpCount = 0 Then
                 ReDim .OpBlock(0)
            Else
                 ReDim .OpBlock(1 To .OpCount)
                 For i = 1 To .OpCount
                       .OpBlock(i).Op = GetWordL()
                       .OpBlock(i).ParaNum = CLng(Val(GetWordL))
                       If .OpBlock(i).ParaNum = 0 Then
                            ReDim .OpBlock(i).Para(0)
                       Else
                            ReDim .OpBlock(i).Para(1 To .OpBlock(i).ParaNum)
                            For m = 1 To .OpBlock(i).ParaNum
                                 .OpBlock(i).Para(m).Value = GetWordL()
                                 BuildQuote_ParamCode .OpBlock(i).Para(m)
                            Next m
                       End If
                 Next i
            End If

        End With

    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ReadTabMatLine", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����ReadMeshLine
'**��    �룺Text(String),OutputMesh(Type_Mesh)
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-01-28 21:24:54
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Sub ReadMeshLine(Text As String, OutputMesh As Type_Mesh)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer
    
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()

        With OutputMesh
        
            .strID = GetWordL()
            .Flags = GetWordL()
            .Resource_Name = GetWordL()
            
            .Translation.X = Val(GetWordL())
            .Translation.Y = Val(GetWordL())
            .Translation.Z = Val(GetWordL())
              
            .Rotation_Angle.X = Val(GetWordL())
            .Rotation_Angle.Y = Val(GetWordL())
            .Rotation_Angle.Z = Val(GetWordL())
            
            .Scale.X = Val(GetWordL())
            .Scale.Y = Val(GetWordL())
            .Scale.Z = Val(GetWordL())
        End With
    
    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModMain", "ReadMeshLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����GetYesNoStr
'**��    �룺lVal(Long) -
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-06 16:38:37
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function GetYesNoStr(ByVal lVal As Long) As String
If lVal = 0 Then
   GetYesNoStr = "��"
Else
   GetYesNoStr = "��"
End If
End Function

'*************************************************************************
'**�� �� ����SaveAll
'**��    �룺 -
'**��    ���� -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-22 15:35:26
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SaveAll()           'A_E
    'Save Quote
     SaveQuote
     
    'Save Items
     SaveItemFile MnBInfo.ModPath & "\item_kinds1.txt"
     SaveItemCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\item_kinds.csv"

    'Save Troops
     SaveTroopFile MnBInfo.ModPath & "\troops.txt"
     SaveTroopCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\troops.csv"
     
    'Save Factions
    SaveFactionFile MnBInfo.ModPath & "\factions.txt"
    SaveFactionCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\factions.csv"
    
    'Save PartyTemplates
    SavePartyTemplateFile MnBInfo.ModPath & "\party_templates.txt"
    SavePartyTemplateCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\party_templates.csv"
    
    'Save Parties
    SavePartyFile MnBInfo.ModPath & "\parties.txt"
    SavePartyCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\parties.csv"
    
    'Save Scenes
    SaveSceneFile MnBInfo.ModPath & "\scenes.txt"
    
    'Save MapIcons
    SaveMapIconFile MnBInfo.ModPath & "\map_icons.txt"
    
    'Save Sound & SoundResource
    SaveSoundFile MnBInfo.ModPath & "\sounds.txt"
    
    'Save Particles System
    SavePSysFile MnBInfo.ModPath & "\particle_systems.txt"
    
    'Save Tableau Materials
    SaveTabMatFile MnBInfo.ModPath & "\tableau_materials.txt"
   
    'Save Meshes
    SaveMeshFile MnBInfo.ModPath & "\meshes.txt"
   
   'Save Triggers
   SaveTriggerFile MnBInfo.ModPath & "\triggers.txt"
   
   'Save Strings
   If IsLoadString Then
      SaveStringFile MnBInfo.ModPath & "\strings.txt"
      SaveStringCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\game_strings.csv"
   End If
   
   'Save
   Dim i As Long
   For i = 1 To UBound(VarNameLists)  '0 global
      SaveVarNameCheckList i
   Next i
   
End Sub

'*************************************************************************
'**�� �� ����ReadAll
'**��    �룺 -
'**��    ���� -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-22 15:42:43
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub ReadAll()   'A_E

frmMain.Caption = Replace(frmMain.Caption, ":[ModName]", ":" & MnBInfo.ModName, , , vbTextCompare)

DoEvents

'ModPic.Picture = LoadPicture(MnBInfo.ModPath & "\Main.bmp")

'Load Sounds
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(9))
LoadSoundFile MnBInfo.ModPath & "\sounds.txt"

'Load Items
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(2))
LoadItemFile MnBInfo.ModPath & "\item_kinds1.txt"
LoadItemCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\item_kinds.csv"

'Load Troops
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(1))
LoadTroopFile MnBInfo.ModPath & "\troops.txt"
LoadTroopCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\troops.csv"

'Load Party_Templates
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(4))
LoadPTFile MnBInfo.ModPath & "\party_templates.txt"
LoadPartyTemplateCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\party_templates.csv"

'Load Parties
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(3))
LoadPartyFile MnBInfo.ModPath & "\parties.txt"
LoadPartyCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\parties.csv"

'Load Factions
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(5))
LoadFactionFile MnBInfo.ModPath & "\factions.txt"
LoadFactionCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\factions.csv"

'Load IModifiers
frmTip.ShowTip ActiveString(PublicTips(1), PublicMsgs(147))
LoadIModFile MnBInfo.MBHome & "\Data\item_modifiers.txt"
LoadIModCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\item_modifiers.csv"

'Load Scenes
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(6))
LoadSceneFile MnBInfo.ModPath & "\scenes.txt"

'Load Particle System
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(8))
LoadPSysFile MnBInfo.ModPath & "\particle_systems.txt"

'Load Map Icons
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(7))
LoadMapIconFile MnBInfo.ModPath & "\map_icons.txt"

'Load Tableau Materials
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(11))
LoadTabMatFile MnBInfo.ModPath & "\tableau_materials.txt"

'Load Meshes
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(12))
LoadMeshFile MnBInfo.ModPath & "\meshes.txt"

'Load Triggers
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(13))
LoadTriggerFile MnBInfo.ModPath & "\triggers.txt"

'Load Global Variables
'frmTip.ShowTip ActiveString(PublicTips(1), PublicTags(Tag_Variable))
LoadGlobalVarFile MnBInfo.ModPath & "\variables.txt"

'Load Quick Strings
frmTip.ShowTip ActiveString(PublicTips(1), PublicTags(Tag_Quick_String))
LoadQuickStringFile MnBInfo.ModPath & "\quick_strings.txt"
LoadQuickStringCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\quick_strings.csv"

'Load Strings
If IsLoadString Then
   frmTip.ShowTip ActiveString(PublicTips(1), PublicTags(Tag_String))
   LoadStringFile MnBInfo.ModPath & "\strings.txt"
   LoadStringCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\game_strings.csv"
End If

'Load Editors
'InitEditorsListView

'Build Quote
frmTip.ShowTip PublicMsgs(148)
BuildQuote

'Read Variable Name Check List
ReadVarNameCheckLists

frmTip.HideTip

frmMain.mSave.Enabled = True
frmMain.mSaveAs.Enabled = True
frmMain.mBackUp.Enabled = True
frmMain.mEditor.Enabled = True
frmMain.mTool.Enabled = True

End Sub

'*************************************************************************
'**�� �� ����MkDirEx
'**��    �룺 (String)FilePath
'**��    ���� -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-12-25 20:16:41
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub MkDirEx(ByVal FilePath As String)
Dim i As Long, k As String, FirstFolder As Boolean, FolderNow As String

FirstFolder = True
For i = 1 To Len(FilePath)
        k = Mid(FilePath, i, 1)
        If k = "\" Then
           If Not FirstFolder Then
              FolderNow = Left(FilePath, i - 1)
              If Dir(FolderNow, vbDirectory + vbHidden + vbNormal + vbSystem) = "" Then
                  MkDir FolderNow
              End If
           Else
              FirstFolder = False
           End If
        End If
Next i

If Dir(FilePath, vbDirectory + vbHidden + vbNormal + vbSystem) = "" Then
      MkDir FilePath
End If

End Sub

'*************************************************************************
'**�� �� ����IsWeapon
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�Ϊ����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-11-29 12:22:10
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsWeapon(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_one_handed_wpn Or IT = itp_type_two_handed_wpn _
Or IT = itp_type_polearm Or IT = itp_type_bow Or IT = itp_type_crossbow Or IT = itp_type_thrown _
Or IT = itp_type_pistol Or IT = itp_type_musket Then
   IsWeapon = True
Else
   IsWeapon = False
End If

End Function

'*************************************************************************
'**�� �� ����IsMeleeWeapon
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�Ϊ��ս����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-3 23:28:30
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsMeleeWeapon(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_one_handed_wpn Or IT = itp_type_two_handed_wpn _
Or IT = itp_type_polearm Then
   IsMeleeWeapon = True
Else
   IsMeleeWeapon = False
End If

End Function

'*************************************************************************
'**�� �� ����IsRangedWeapon
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�ΪԶ������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-3 23:29:35
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsRangedWeapon(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_bow Or IT = itp_type_crossbow Or IT = itp_type_thrown Then
   IsRangedWeapon = True
Else
   IsRangedWeapon = False
End If

End Function

'*************************************************************************
'**�� �� ����IsFireArm
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�Ϊ����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-3 23:29:40
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsFireArm(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_pistol Or IT = itp_type_musket Then
   IsFireArm = True
Else
   IsFireArm = False
End If

End Function

'*************************************************************************
'**�� �� ����IsAmmo
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�Ϊ��ҩ
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-3 23:29:45
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsAmmo(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_arrows Or IT = itp_type_bolts Or IT = itp_type_bullets Then
   IsAmmo = True
Else
   IsAmmo = False
End If

End Function

'*************************************************************************
'**�� �� ����IsHorse
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�Ϊ��ƥ
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-07 08:52:25
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsHorse(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_horse Then
   IsHorse = True
Else
   IsHorse = False
End If

End Function

'*************************************************************************
'**�� �� ����IsShield
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�Ϊ����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-07 08:52:25
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsShield(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_shield Then
   IsShield = True
Else
   IsShield = False
End If

End Function

'*************************************************************************
'**�� �� ����IsGood
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�Ϊ����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-08 23:15:55
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsGood(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_goods Then
   IsGood = True
Else
   IsGood = False
End If

End Function

'*************************************************************************
'**�� �� ����IsFood
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�Ϊʳ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-08 23:16:22
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsFood(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)

If ChkBit64b(tI, itp_food) Then
   IsFood = True
Else
   IsFood = False
End If

End Function

'*************************************************************************
'**�� �� ����IsBook
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�Ϊ�鱾
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-08 23:16:00
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsBook(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_book Then
   IsBook = True
Else
   IsBook = False
End If

End Function


'*************************************************************************
'**�� �� ����IsAnimal
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�Ϊ����
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-08 23:16:05
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsAnimal(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_animal Then
   IsAnimal = True
Else
   IsAnimal = False
End If

End Function
'*************************************************************************
'**�� �� ����IsArmor
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�Ϊ������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-07 14:48:42
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsArmor(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_head_armor Or IT = itp_type_hand_armor Or IT = itp_type_body_armor Or IT = itp_type_foot_armor Then
   IsArmor = True
Else
   IsArmor = False
End If

End Function
'*************************************************************************
'**�� �� ����IsHeadArmor
'**��    �룺(Str)ItmType
'**��    ������
'**�����������ж���Ʒ�Ƿ�Ϊͷ��������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-21 15:18:22
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsHeadArmor(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_head_armor Then
   IsHeadArmor = True
Else
   IsHeadArmor = False
End If

End Function

'*************************************************************************
'**�� �� ����GetItmType
'**��    �룺(Str)ItmType
'**��    ����(Integer)-
'**�����������õ���������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-11-29 22:24:34
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function GetItmType(itmType As String) As Integer

If Val(itmType) < 2100000000 Then

    GetItmType = Val(itmType) Mod 256
Else

    GetItmType = MinusIT(itmType) Mod 256
End If

End Function

'*************************************************************************
'**�� �� ����HaveDoubleUsages
'**��    �룺(Str)ItmType
'**��    ����(Boolean)-
'**�����������ж���Ʒ�Ƿ�������ʹ��ģʽ
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-11-29 22:10:51
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function HaveDoubleUsages(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
HaveDoubleUsages = ChkBit64b(tI, itp_next_item_as_melee)

End Function

'*************************************************************************
'**�� �� ����IsTwoHanded
'**��    �룺(Str)ItmType
'**��    ����(Boolean)-
'**�����������ж���Ʒ�Ƿ���Գֶ�
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-4-12 23:19:34
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsTwoHanded(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsTwoHanded = ChkBit64b(tI, itp_two_handed)

End Function

'*************************************************************************
'**�� �� ����IsBonusAgainstShield
'**��    �룺(Str)ItmType
'**��    ����(Boolean)-
'**�����������ж���Ʒ�Ƿ�Զܼӳ�
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-4-13 21:51:12
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsBonusAgainstShield(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsBonusAgainstShield = ChkBit64b(tI, itp_bonus_against_shield)

End Function

'*************************************************************************
'**�� �� ����IsUnbalanced
'**��    �룺(Str)ItmType
'**��    ����(Boolean)-
'**�����������ж���Ʒ�Ƿ�ƽ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-4-13 22:31:46
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsUnbalanced(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsUnbalanced = ChkBit64b(tI, itp_unbalanced)

End Function

'*************************************************************************
'**�� �� ����IsCrushThrough
'**��    �룺(Str)ItmType
'**��    ����(Boolean)-
'**�����������ж���Ʒ�Ƿ��Ƹ�
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-4-13 22:33:34
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsCrushThrough(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsCrushThrough = ChkBit64b(tI, itp_crush_through)

End Function

'*************************************************************************
'**�� �� ����IsCantUseOnHorseBack
'**��    �룺(Str)ItmType
'**��    ����(Boolean)-
'**�����������ж���Ʒ�Ƿ���������ʹ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-4-13 22:46:11
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsCantUseOnHorseBack(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsCantUseOnHorseBack = ChkBit64b(tI, itp_cant_use_on_horseback)

End Function
'*************************************************************************
'**�� �� ����IsCantReloadOnHorseBack
'**��    �룺(Str)ItmType
'**��    ����(Boolean)-
'**�����������ж���Ʒ�Ƿ���������װ��
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-4-13 22:54:22
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsCantReloadOnHorseBack(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsCantReloadOnHorseBack = ChkBit64b(tI, itp_cant_reload_on_horseback)

End Function

'*************************************************************************
'**�� �� ����IsCanPenetrateShield
'**��    �룺(Str)ItmType
'**��    ����(Boolean)-
'**�����������ж���Ʒ�Ƿ��ܴ���
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-4-13 22:50:34
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function IsCanPenetrateShield(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsCanPenetrateShield = ChkBit64b(tI, itp_can_penetrate_shield)

End Function
'*************************************************************************
'**�� �� ����GetAttachment
'**��    �룺(String)itmType
'**��    ����(integer)-
'**�����������õ�������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-11-29 22:24:34
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function GetAttachment(itmType As String) As Integer
Dim tI As Integer64b, tI2 As Integer64b

tI = StrToI64(itmType)
GetAttachment = 0

If ChkBit64b(tI, itp_Attachment_Left_bit) And ChkBit64b(tI, itp_Attachment_Right_bit) And _
ChkBit64b(tI, itp_Attachment_Armature_bit1) And ChkBit64b(tI, itp_Attachment_Armature_bit2) Then
    GetAttachment = 4
    Exit Function
ElseIf ChkBit64b(tI, itp_Attachment_Left_bit) And ChkBit64b(tI, itp_Attachment_Right_bit) Then
    GetAttachment = 3
    Exit Function
ElseIf ChkBit64b(tI, itp_Attachment_Left_bit) Then
    GetAttachment = 1
    Exit Function
ElseIf ChkBit64b(tI, itp_Attachment_Right_bit) Then
    GetAttachment = 2
    Exit Function
End If

End Function

'*************************************************************************
'**�� �� ����GetDamage
'**��    �룺(Long)Damage,(Integer)Damage_Type
'**��    ����(Long)-
'**�����������õ��˺���ֵ������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-11-29 22:24:34
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function GetDamage(Damage As Long, Damage_Type As Integer) As Long

If Damage - 2 ^ 9 >= 0 Then
    Damage_Type = 2
    GetDamage = Damage - 2 ^ 9
ElseIf Damage - 2 ^ 8 >= 0 Then
    Damage_Type = 1
    GetDamage = Damage - 2 ^ 8
Else
    Damage_Type = 0
    GetDamage = Damage
End If

End Function

'*************************************************************************
'**�� �� ����ExDamage
'**��    �룺(Long)Damage,(Integer)Damage_Type
'**��    ����(Long)-
'**�����������õ�TXT�м�¼��ʽ���˺���ֵ
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2010-12-07 22:10:38
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function ExDamage(Damage As Long, Damage_Type As Integer) As Long

If Damage_Type > 0 Then
ExDamage = Damage + 2 ^ (7 + Damage_Type)
ElseIf Damage_Type = 0 Then
ExDamage = Damage
End If

End Function

'*************************************************************************
'**�� �� ����strIIf
'**��    �룺(Boolean)Condition,(Boolean)strTrue,(Boolean)strFalse
'**��    ����(String)-
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-08 21:56:24
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function strIIf(Condition As Boolean, strTrue As String, strFalse As String) As String

If Condition Then
   strIIf = strTrue
Else
   strIIf = strFalse
End If
End Function

'*************************************************************************
'**�� �� ����ReadItemCSVLine
'**��    �룺(String)Text
'**��    ����(Long)-
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-25 13:41:24
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function ReadItemCSVLine(ByVal Text As String) As Long
Dim i As Integer, n As Integer, TemP As String, StrLine() As String, arrTmp() As String, j As Integer

StrLine() = Split(Text, vbCrLf)
    For i = 0 To UBound(StrLine)
        TemP = StrLine(i)
        If Len(Trim$(TemP)) < 1 Then
            TemP = ""
        Else
            arrTmp = Split(TemP, "|")

            For n = 0 To TemItemCount - 1
                If LCase$(TemItems(n).dbName) = LCase$(arrTmp(0)) Then
                    TemItems(n).csvName = arrTmp(1)
                    'Exit For
                    j = j + 1
                End If
            
            
            If Right(arrTmp(0), 3) = "_pl" Then
                     If LCase$(TemItems(n).dbName) = LCase$(Left(arrTmp(0), Len(arrTmp(0)) - 3)) Then
                        TemItems(n).csvName_pl = arrTmp(1)
                        'Exit For
                     End If
            End If
            
            Next n
        End If
    Next i
    
ReadItemCSVLine = j
End Function

'*************************************************************************
'**�� �� ����LTrimEx
'**��    �룺(String)Text
'**��    ����(String)-
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-25 13:53:30
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function LTrimEx(ByVal Text As String) As String
Dim i As Integer, k As Long, temStr As String, n As Integer

If Len(Text) > 0 Then
For i = 1 To Len(Text)
     k = Asc(Mid(Text, i, 1))
     
     If k <> 13 And k <> 10 And k <> 32 Then
        n = i
     End If
     
Next i

If n > 0 Then
   LTrimEx = Right(Text, Len(Text) - n + 1)
Else
   
End If

End If
End Function

Public Function GetEditorIndex(ByVal Tag As String) As Long
Dim s As String

s = Right(Tag, Len(Tag) - 5)
GetEditorIndex = Val(s)
End Function

Public Function ActiveString(ByVal StrMain As String, ParamArray StrNew() As Variant) As String
Dim sKey As String, i As Integer, strTem As String

strTem = StrMain

For i = 0 To UBound(StrNew)
  sKey = "[str" & i & "]"
  strTem = Replace(strTem, sKey, CStr(StrNew(i)), , , vbTextCompare)
  
Next i

ActiveString = strTem
End Function

'*************************************************************************
'**�� �� ����StructureTroopInventory
'**��    �룺(Long)Trp_Idx
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-01-19 20:05:16
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub StructureTroopInventory(Optional Trp_Idx As Long = -1)
Dim i As Integer, TemList() As Type_XY_Index, TrpVes As Type_Troops

If Trp_Idx <= -1 Then
   TrpVes = CurrentTrp
Else
   TrpVes = Trps(Trp_Idx)
End If

With TrpVes
ReDim TemList(0)
For i = 1 To 64
        If .lstInventory(i).X > -1 Then
            ReDim Preserve TemList(UBound(TemList) + 1)
            TemList(UBound(TemList)) = .lstInventory(i)
        End If
Next i

   For i = 1 To 64
     If i <= UBound(TemList) Then
           .lstInventory(i) = TemList(i)
     Else
           .lstInventory(i).X = -1
           .lstInventory(i).strX = ""
           .lstInventory(i).Y = 0
     End If
   Next i

End With

If Trp_Idx <= -1 Then
   CurrentTrp = TrpVes
Else
    Trps(Trp_Idx) = TrpVes
End If

End Sub

'*************************************************************************
'**�� �� ����StructurePartyStacksEx
'**��    �룺(Long)Party_Idx
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-03-20 11:59:01
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub StructurePartyStacksEx(ByVal Party_Idx As Long)
Dim i As Integer, TemList() As Type_Stacks

With Parties(Party_Idx)
ReDim TemList(0)
For i = 1 To .StacksCount
        If .Stacks(i).ID > -1 Then
            ReDim Preserve TemList(UBound(TemList) + 1)
            TemList(UBound(TemList)) = .Stacks(i)
        End If
Next i

.StacksCount = UBound(TemList)

If .StacksCount > 0 Then
ReDim .Stacks(1 To .StacksCount)
   For i = 1 To .StacksCount
           .Stacks(i) = TemList(i)
   Next i
End If
End With
End Sub

'*************************************************************************
'**�� �� ����StructureItemFactions
'**��    �룺(Long)Itm_Idx
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-03-21 17:07:25
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub StructureItemFactions(ByVal Itm_Idx As Long)
Dim i As Integer, TemList() As Type_Chest, ItmVes As Type_Item

If Itm_Idx <= -1 Then
   ItmVes = CurrentItm
Else
   ItmVes = itm(Itm_Idx)
End If

With ItmVes
ReDim TemList(0)
For i = 1 To .FactionCount
        If .Faction(i).ID > -1 Then
            ReDim Preserve TemList(UBound(TemList) + 1)
            TemList(UBound(TemList)) = .Faction(i)
        End If
Next i

.FactionCount = UBound(TemList)

If .FactionCount > 0 Then
ReDim .Faction(1 To .FactionCount)
     For i = 1 To .FactionCount
        .Faction(i) = TemList(i)
     Next i
End If
End With

If Itm_Idx <= -1 Then
   CurrentItm = ItmVes
Else
   itm(Itm_Idx) = ItmVes
End If
End Sub

'*************************************************************************
'**�� �� ����StructureSceneChests
'**��    �룺(Long)Scene_Idx
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-03-20 22:59:44
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub StructureSceneChests(ByVal Scene_Idx As Long)
Dim i As Integer, TemList() As Type_Chest, SceneVes As Type_Scene

If Scene_Idx <= -1 Then
   SceneVes = CurrentScene
Else
   SceneVes = Scenes(Scene_Idx)
End If

With SceneVes
ReDim TemList(0)
For i = 1 To .ChestCount
        If .Chests(i).ID > -1 Then
            ReDim Preserve TemList(UBound(TemList) + 1)
            TemList(UBound(TemList)) = .Chests(i)
        End If
Next i

.ChestCount = UBound(TemList)

If .ChestCount > 0 Then
ReDim .Chests(1 To .ChestCount)
     For i = 1 To .ChestCount
       .Chests(i) = TemList(i)
     Next i
End If
End With

If Scene_Idx <= -1 Then
   CurrentScene = SceneVes
Else
   Scenes(Scene_Idx) = SceneVes
End If
End Sub

'*************************************************************************
'**�� �� ����StructureSoundRes
'**��    �룺(Long)Snd_Idx
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-03-20 23:25:33
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub StructureSoundRes(ByVal Snd_Idx As Long)
Dim i As Integer, TemList() As Type_ResourceInSound, SoundVes As Type_Sound

If Snd_Idx <= -1 Then
   SoundVes = CurrentSound
Else
   SoundVes = Sounds(Snd_Idx)
End If

With SoundVes
ReDim TemList(0)
For i = 1 To .ResourceCount
        If .Resource(i).ID > -1 Then
            ReDim Preserve TemList(UBound(TemList) + 1)
            TemList(UBound(TemList)) = .Resource(i)
        End If
Next i

.ResourceCount = UBound(TemList)

If .ResourceCount > 0 Then
ReDim .Resource(1 To .ResourceCount)
     For i = 1 To .ResourceCount
      .Resource(i) = TemList(i)
     Next i
End If
End With

If Snd_Idx <= -1 Then
   CurrentSound = SoundVes
Else
   Sounds(Snd_Idx) = SoundVes
End If
End Sub


'*************************************************************************
'**�� �� ����StructureFactionRelationShips
'**��    �룺
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-03-20 13:14:51
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub StructureFactionRelationShips(ByVal Faction_Idx As Long)
Dim i As Long, FacVes As Type_Faction, TemRels() As Type_RelationShip

If Faction_Idx > -1 Then
  FacVes = Factions(Faction_Idx)
Else
  FacVes = CurrentFaction
End If

ReDim TemRels(N_Faction - 1)

With FacVes
     For i = 0 To UBound(.RelationShip)
         .RelationShip(i).ID = GetID(.RelationShip(i).strID, False, "", -1)
         
         If .RelationShip(i).ID > -1 Then
            TemRels(.RelationShip(i).ID) = .RelationShip(i)
         End If
     Next i
     
     ReDim .RelationShip(N_Faction - 1)
     .RelationShip = TemRels
End With


If Faction_Idx > -1 Then
  Factions(Faction_Idx) = FacVes
Else
  CurrentFaction = FacVes
End If

End Sub

Public Sub PurseSceneInTroop(ByVal HexScene As String, SceneID As Long, Scene_strID As String, Entry As Long)
Dim strTem As String, lngTem As Long

       strTem = "0000" & Hex(HexScene)
       lngTem = Val("&H" & Right(strTem, 4))
       SceneID = lngTem
       Scene_strID = Scenes(lngTem).strID
       Entry = Val("&H" & Left(strTem, Len(strTem) - 4))
         
End Sub

'*************************************************************************
'**�� �� ����EnumEntities
'**��    �룺(Integer)Tag
'**��    ����(Boolean)
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-07-09 23:22:14
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function EnumEntities(Tag As Integer, strModule As String, oStr() As String) As Boolean
Dim i As Long

Select Case Tag
     Case 0
       EnumEntities = False
     Case Tag_Register
       EnumEntities = False
     Case Tag_Variable
       If UBound(TemGVarNameList.Triggers) >= 1 Then
         ReDim oStr(UBound(TemGVarNameList.Triggers(1).Checks))
         For i = 0 To UBound(TemGVarNameList.Triggers(1).Checks)
           With TemGVarNameList.Triggers(1)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .Checks(i))
             oStr(i) = Replace(oStr(i), "[csvname]", PYTags(Tag) & .Checks(i))
             oStr(i) = Replace(oStr(i), "[csvname_pl]", PYTags(Tag) & .Checks(i))
             oStr(i) = Replace(oStr(i), "[disname]", PYTags(Tag) & .Checks(i))
             oStr(i) = Replace(oStr(i), "[disname_pl]", PYTags(Tag) & .Checks(i))
             'oStr(i) = Replace(oStr(i), "[value]", PYTags(Tag) & .Checks(i))
           End With
         Next i
         EnumEntities = True
       Else
         EnumEntities = False
       End If
         
     Case Tag_String
       EnumEntities = False   'A-E
     Case Tag_Item
       ReDim oStr(N_Item - 1)
       For i = 0 To N_Item - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", itm(i).dbName)
         oStr(i) = Replace(oStr(i), "[csvname]", itm(i).csvName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", itm(i).csvName_pl)
         oStr(i) = Replace(oStr(i), "[disname]", itm(i).disname)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Item, i))
       Next i
       EnumEntities = True
     Case Tag_Troop
       ReDim oStr(N_Troop - 1)
       For i = 0 To N_Troop - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", Trps(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", Trps(i).csvName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", Trps(i).csvName_pl)
         oStr(i) = Replace(oStr(i), "[disname]", Trps(i).strName)
         oStr(i) = Replace(oStr(i), "[disname_pl]", Trps(i).strPtName)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Troop, i))
       Next i
       EnumEntities = True
     Case Tag_Faction
       ReDim oStr(N_Faction - 1)
       For i = 0 To N_Faction - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", Factions(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", Factions(i).csvName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", Factions(i).csvName)
         oStr(i) = Replace(oStr(i), "[disname]", Factions(i).strName)
         oStr(i) = Replace(oStr(i), "[disname_pl]", Factions(i).strName)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Faction, i))
       Next i
       EnumEntities = True
     Case Tag_Quest
       EnumEntities = False   'A-E
     Case Tag_Party_Tpl
       ReDim oStr(N_PT - 1)
       For i = 0 To N_PT - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", PTs(i).ptID)
         oStr(i) = Replace(oStr(i), "[csvname]", PTs(i).csvName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", PTs(i).csvName)
         oStr(i) = Replace(oStr(i), "[disname]", PTs(i).ptName)
         oStr(i) = Replace(oStr(i), "[disname_pl]", PTs(i).ptName)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Party_Tpl, i))
       Next i
       EnumEntities = True
     Case Tag_Party
       ReDim oStr(N_Party - 1)
       For i = 0 To N_Party - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", Parties(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", Parties(i).csvName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", Parties(i).csvName)
         oStr(i) = Replace(oStr(i), "[disname]", Parties(i).strName)
         oStr(i) = Replace(oStr(i), "[disname_pl]", Parties(i).strName)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Party, i))
       Next i
       EnumEntities = True
     Case Tag_Scene
       ReDim oStr(N_Scene - 1)
       For i = 0 To N_Scene - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", Scenes(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", Scenes(i).strName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", Scenes(i).strName)
         oStr(i) = Replace(oStr(i), "[disname]", Scenes(i).strName)
         oStr(i) = Replace(oStr(i), "[disname_pl]", Scenes(i).strName)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Scene, i))
       Next i
       EnumEntities = True
     Case Tag_Mission_tpl
       EnumEntities = False   'A-E
     Case Tag_Menu
       EnumEntities = False   'A-E
     Case Tag_Script
       EnumEntities = False   'A-E
     Case Tag_Particle_Sys
       ReDim oStr(N_PSys - 1)
       For i = 0 To N_PSys - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", PSys(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", PSys(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", PSys(i).strID)
         oStr(i) = Replace(oStr(i), "[disname]", PSys(i).strID)
         oStr(i) = Replace(oStr(i), "[disname_pl]", PSys(i).strID)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Particle_Sys, i))
       Next i
       EnumEntities = True
     Case Tag_Scene_Prop
       EnumEntities = False   'A-E
     Case Tag_Sound
       ReDim oStr(N_Sound - 1)
       For i = 0 To N_Sound - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", Sounds(i).sndName)
         oStr(i) = Replace(oStr(i), "[csvname]", Sounds(i).sndName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", Sounds(i).sndName)
         oStr(i) = Replace(oStr(i), "[disname]", Sounds(i).sndName)
         oStr(i) = Replace(oStr(i), "[disname_pl]", Sounds(i).sndName)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Sound, i))
       Next i
       EnumEntities = True
     Case Tag_Local_Variable
       If UBound(CurVarNameList.Triggers) >= CheckListTrgIdx Then
         ReDim oStr(UBound(CurVarNameList.Triggers(CheckListTrgIdx).Checks))
         For i = 0 To UBound(CurVarNameList.Triggers(CheckListTrgIdx).Checks)
           With CurVarNameList.Triggers(CheckListTrgIdx)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .Checks(i))
             oStr(i) = Replace(oStr(i), "[csvname]", PYTags(Tag) & .Checks(i))
             oStr(i) = Replace(oStr(i), "[csvname_pl]", PYTags(Tag) & .Checks(i))
             oStr(i) = Replace(oStr(i), "[disname]", PYTags(Tag) & .Checks(i))
             oStr(i) = Replace(oStr(i), "[disname_pl]", PYTags(Tag) & .Checks(i))
             'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Local_Variable, i))
           End With
         Next i
         EnumEntities = True
       Else
         EnumEntities = False
       End If
     Case Tag_Map_Icon
       ReDim oStr(N_MapIcon - 1)
       For i = 0 To N_MapIcon - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", MapIcons(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", MapIcons(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", MapIcons(i).strID)
         oStr(i) = Replace(oStr(i), "[disname]", MapIcons(i).strID)
         oStr(i) = Replace(oStr(i), "[disname_pl]", MapIcons(i).strID)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Map_Icon, i))
       Next i
       EnumEntities = True
     Case Tag_Skill
       ReDim oStr(UBound(PublicSkills))
       For i = 0 To UBound(PublicSkills)
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", PublicSkills(i))
         oStr(i) = Replace(oStr(i), "[csvname]", PublicSkills(i))
         oStr(i) = Replace(oStr(i), "[csvname_pl]", PublicSkills(i))
         oStr(i) = Replace(oStr(i), "[disname]", PublicSkills(i))
         oStr(i) = Replace(oStr(i), "[disname_pl]", PublicSkills(i))
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Skill, i))
       Next i
       EnumEntities = True
     Case Tag_Mesh
       ReDim oStr(N_Mesh - 1)
       For i = 0 To N_Mesh - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", Mesh(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", Mesh(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", Mesh(i).strID)
         oStr(i) = Replace(oStr(i), "[disname]", Mesh(i).strID)
         oStr(i) = Replace(oStr(i), "[disname_pl]", Mesh(i).strID)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Mesh, i))
       Next i
       EnumEntities = True
     Case Tag_Presentation
       EnumEntities = False   'A-E
     Case Tag_Quick_String
       EnumEntities = False   'A-E
     Case Tag_Track
       EnumEntities = False   'A-E
     Case Tag_Tableau
       ReDim oStr(N_TabMat - 1)
       For i = 0 To N_TabMat - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", TabMat(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", TabMat(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", TabMat(i).strID)
         oStr(i) = Replace(oStr(i), "[disname]", TabMat(i).strID)
         oStr(i) = Replace(oStr(i), "[disname_pl]", TabMat(i).strID)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Tableau, i))
       Next i
       EnumEntities = True
     Case Tag_Animation
       EnumEntities = False   'A-E
     Case Tags_End
       EnumEntities = False
   End Select
End Function


'*************************************************************************
'**�� �� ����GetEntityName
'**��    �룺(Integer)Tag
'**��    ����(String)
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-07-09 23:22:14
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function GetEntityName(Tag As Integer, Index As Long, strModule As String) As String
Dim i As Long
i = Index
Select Case Tag
     Case 0
       GetEntityName = ""
     Case Tag_Register
       GetEntityName = ""
     Case Tag_Variable
       GetEntityName = ""
     Case Tag_String
       GetEntityName = ""   'A-E
     Case Tag_Item
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[dbname]", itm(i).dbName)
         GetEntityName = Replace(GetEntityName, "[csvname]", itm(i).csvName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", itm(i).csvName_pl)
         GetEntityName = Replace(GetEntityName, "[disname]", itm(i).disname)

     Case Tag_Troop
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", Trps(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", Trps(i).csvName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", Trps(i).csvName_pl)
         GetEntityName = Replace(GetEntityName, "[disname]", Trps(i).strName)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", Trps(i).strPtName)

     Case Tag_Faction
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", Factions(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", Factions(i).csvName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", Factions(i).csvName)
         GetEntityName = Replace(GetEntityName, "[disname]", Factions(i).strName)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", Factions(i).strName)

     Case Tag_Quest
       GetEntityName = False   'A-E
     Case Tag_Party_Tpl
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", PTs(i).ptID)
         GetEntityName = Replace(GetEntityName, "[csvname]", PTs(i).csvName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", PTs(i).csvName)
         GetEntityName = Replace(GetEntityName, "[disname]", PTs(i).ptName)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", PTs(i).ptName)

     Case Tag_Party
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", Parties(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", Parties(i).csvName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", Parties(i).csvName)
         GetEntityName = Replace(GetEntityName, "[disname]", Parties(i).strName)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", Parties(i).strName)

     Case Tag_Scene
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", Scenes(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", Scenes(i).strName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", Scenes(i).strName)
         GetEntityName = Replace(GetEntityName, "[disname]", Scenes(i).strName)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", Scenes(i).strName)

     Case Tag_Mission_tpl
       GetEntityName = ""   'A-E
     Case Tag_Menu
       GetEntityName = ""   'A-E
     Case Tag_Script
       GetEntityName = ""   'A-E
     Case Tag_Particle_Sys
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", PSys(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", PSys(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", PSys(i).strID)
         GetEntityName = Replace(GetEntityName, "[disname]", PSys(i).strID)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", PSys(i).strID)

     Case Tag_Scene_Prop
       GetEntityName = ""   'A-E
     Case Tag_Sound
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", Sounds(i).sndName)
         GetEntityName = Replace(GetEntityName, "[csvname]", Sounds(i).sndName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", Sounds(i).sndName)
         GetEntityName = Replace(GetEntityName, "[disname]", Sounds(i).sndName)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", Sounds(i).sndName)

     Case Tag_Local_Variable
       GetEntityName = ""
     Case Tag_Map_Icon
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", MapIcons(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", MapIcons(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", MapIcons(i).strID)
         GetEntityName = Replace(GetEntityName, "[disname]", MapIcons(i).strID)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", MapIcons(i).strID)

     Case Tag_Skill
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", PublicSkills(i))
         GetEntityName = Replace(GetEntityName, "[csvname]", PublicSkills(i))
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", PublicSkills(i))
         GetEntityName = Replace(GetEntityName, "[disname]", PublicSkills(i))
         GetEntityName = Replace(GetEntityName, "[disname_pl]", PublicSkills(i))

     Case Tag_Mesh
       GetEntityName = ""   'A-E
     Case Tag_Presentation
       GetEntityName = ""   'A-E
     Case Tag_Quick_String
       GetEntityName = ""   'A-E
     Case Tag_Track
       GetEntityName = ""   'A-E
     Case Tag_Tableau
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", TabMat(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", TabMat(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", TabMat(i).strID)
         GetEntityName = Replace(GetEntityName, "[disname]", TabMat(i).strID)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", TabMat(i).strID)

     Case Tag_Animation
       GetEntityName = ""   'A-E
     Case Tags_End
       GetEntityName = ""
   End Select
End Function

'*************************************************************************
'**�� �� ����Max_Int
'**��    �룺(Int)a,(Int)b
'**��    ����(Integer) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-07-20 20:49:52
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function Max_Int(a As Integer, b As Integer) As Integer
Max_Int = IIf(a >= b, a, b)
End Function


'*************************************************************************
'**�� �� ����EnumConsts
'**��    �룺(String)ParaType
'**��    ����(Boolean)
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2011-07-09 23:22:14
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Function EnumConsts(ParaType As String, strModule As String, oStr() As String) As Boolean
Dim i As Long

Select Case ParaType     'ends_add
     Case ""
       EnumConsts = False
     Case "pos"
       EnumConsts = False
     Case "s"
       EnumConsts = False
     Case "itp"
         ReDim oStr(1 To UBound(Item_Type))
         For i = 1 To UBound(Item_Type)
           With Item_Type(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .X)
             oStr(i) = Replace(oStr(i), "[csvname]", .Y)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .Y)
             oStr(i) = Replace(oStr(i), "[disname]", .X)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .X)
             oStr(i) = Replace(oStr(i), "[value]", i)
           End With
         Next i
         EnumConsts = True
     Case "tf"
         ReDim oStr(0 To UBound(Tf))
         For i = 0 To UBound(Tf)
           With Tf(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .strName)
             oStr(i) = Replace(oStr(i), "[csvname]", .csvName)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .csvName)
             oStr(i) = Replace(oStr(i), "[disname]", .strName)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .strName)
             oStr(i) = Replace(oStr(i), "[value]", I64toStrNZ(.Value))
           End With
         Next i
         EnumConsts = True
     Case "pf"
         ReDim oStr(0 To UBound(Pf))
         For i = 0 To UBound(Pf)
           With Pf(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .strName)
             oStr(i) = Replace(oStr(i), "[csvname]", .csvName)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .csvName)
             oStr(i) = Replace(oStr(i), "[disname]", .strName)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .strName)
             oStr(i) = Replace(oStr(i), "[value]", I64toStrNZ(.Value))
           End With
         Next i
         EnumConsts = True
     Case "bs"
         ReDim oStr(0 To UBound(BoolSwitch))
         For i = 0 To UBound(BoolSwitch)
           With BoolSwitch(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .X)
             oStr(i) = Replace(oStr(i), "[csvname]", .Y)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .Y)
             oStr(i) = Replace(oStr(i), "[disname]", .X)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .X)
             oStr(i) = Replace(oStr(i), "[value]", i)
           End With
         Next i
         EnumConsts = True
     Case "ap"
         ReDim oStr(0 To UBound(AccessPrivilege))
         For i = 0 To UBound(AccessPrivilege)
           With AccessPrivilege(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .X)
             oStr(i) = Replace(oStr(i), "[csvname]", .Y)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .Y)
             oStr(i) = Replace(oStr(i), "[disname]", .X)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .X)
             oStr(i) = Replace(oStr(i), "[value]", i)
           End With
         Next i
         EnumConsts = True
     Case "as"
         ReDim oStr(0 To UBound(AbsSwitch))
         For i = 0 To UBound(AbsSwitch)
           With AbsSwitch(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .X)
             oStr(i) = Replace(oStr(i), "[csvname]", .Y)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .Y)
             oStr(i) = Replace(oStr(i), "[disname]", .X)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .X)
             oStr(i) = Replace(oStr(i), "[value]", i)
           End With
         Next i
         EnumConsts = True
     Case "ai_bhvr"
         ReDim oStr(0 To UBound(AI_Bhvr))
         For i = 0 To UBound(AI_Bhvr)
           With AI_Bhvr(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .X)
             oStr(i) = Replace(oStr(i), "[csvname]", .Y)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .Y)
             oStr(i) = Replace(oStr(i), "[disname]", .X)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .X)
             oStr(i) = Replace(oStr(i), "[value]", i)
           End With
         Next i
         EnumConsts = True
     Case "po"
         ReDim oStr(0 To UBound(PlayOption))
         For i = 0 To UBound(PlayOption)
           With PlayOption(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .X)
             oStr(i) = Replace(oStr(i), "[csvname]", .Y)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .Y)
             oStr(i) = Replace(oStr(i), "[disname]", .X)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .X)
             oStr(i) = Replace(oStr(i), "[value]", i)
           End With
         Next i
         EnumConsts = True
    Case Else
      EnumConsts = False
   End Select
End Function
'*************************************************************************
'**�� �� ����CancelTopForms
'**��    �룺
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-03-05 11:29:25
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub CancelTopForms()
  If IsLoadString Then
     SetWindowPos frmStrTool.hWnd, -2, 0, 0, 0, 0, 3
  End If
End Sub
'*************************************************************************
'**�� �� ����SetTopForms
'**��    �룺
'**��    ����
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-03-05 11:31:35
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************
Public Sub SetTopForms()
  If IsLoadString Then
     SetWindowPos frmStrTool.hWnd, -1, 0, 0, 0, 0, 3
  End If
End Sub
