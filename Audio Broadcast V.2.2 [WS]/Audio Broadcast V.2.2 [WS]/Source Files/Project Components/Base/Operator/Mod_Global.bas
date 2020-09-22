Attribute VB_Name = "Mod_Global"
Option Explicit
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __              ___            ___          |###'
'###|        | ____ \              |  |             \  \          /  /          |###'
'###|        | |   \ \             |  |              \  \        /  /           |###'
'###|        | |    \ \            |  |               \  \      /  /            |###'
'###|        | |    / /            |  |                \  \    /  /             |###'
'###|        | |___/ /     __      |  |______    __     \  \__/  /              |###'
'###|        |______/     (__)     |_________|  (__)     \______/               |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|  Global Variables   |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|

'###|
'###|
'###|
'###|
'###|
'###|
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __                   ________               |###'
'###|        | ____ \              |  |                 |  ______|              |###'
'###|        | |   \ \             |  |                 | |                     |###'
'###|        | |    \ \            |  |                 | |______               |###'
'###|        | |    / /            |  |                 |  ______|              |###'
'###|        | |___/ /     __      |  |______    __     | |______               |###'
'###|        |______/     (__)     |_________|  (__)    |________|              |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|    Global Enum      |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|

'###|
'###|
'###|
'###|
'###|
'###|
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __                   ________               |###'
'###|        | ____ \              |  |                 |  ______)              |###'
'###|        | |   \ \             |  |                 | |                     |###'
'###|        | |    \ \            |  |                 | |                     |###'
'###|        | |    / /            |  |                 | |                     |###'
'###|        | |___/ /     __      |  |______    __     | |______               |###'
'###|        |______/     (__)     |_________|  (__)    |________)              |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|    Global Const     |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|

'###|
'###|
'###|
'###|
'###|
'###|
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __                  ______________          |###'
'###|        | ____ \              |  |                |_____    _____|         |###'
'###|        | |   \ \             |  |                      |  |               |###'
'###|        | |    \ \            |  |                      |  |               |###'
'###|        | |    / /            |  |                      |  |               |###'
'###|        | |___/ /     __      |  |______    __          |  |               |###'
'###|        |______/     (__)     |_________|  (__)         |__|               |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|     Global Type     |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|
        Type TagInitCommonControlsEx
            LngSize                                 As Long
            LngICC                                  As Long
        End Type
'###|
'###|
'###|
'###|
'###|
'###|
'###|###########################################################################|###'
'###|###########################################################################|###'
'###|                                                                           |###'
'###|         ______                __                      ______              |###'
'###|        | ____ \              |  |                    /  __  \             |###'
'###|        | |   \ \             |  |                   /  /  \  \            |###'
'###|        | |    \ \            |  |                  /  /____\  \           |###'
'###|        | |    / /            |  |                 /  /______\  \          |###'
'###|        | |___/ /     __      |  |______    __    /  /        \  \         |###'
'###|        |______/     (__)     |_________|  (__)  /__/          \__\        |###'
'###|                                                                           |###'
'###|                                                                           |###'
'###|###########################################################################|###'
'###|#######################|#|########################|#|######################|###'
'###|#|                     |#|                        |#|                      |###'
'###|#|    Global API       |#|                        |#|                      |###'
'###|#|_____________________|#|________________________|#|______________________|###'
'###|#######################|#|########################|#|######################|###'
'###|
        Declare Function InitCommonControlsEx Lib _
                                "comctl32.dll" ( _
                                iccex As TagInitCommonControlsEx) As _
                                Boolean