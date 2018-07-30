Attribute VB_Name = "OLD_Draw"
Private Sub DrawGDI()
    'Cycle Through in-game stuff before cycling through editors
    If frmMenu.visible Then
        If frmMenu.picCharacter.visible Then NewCharacterDrawSprite
    End If
    
    If frmMain.visible Then
        If frmMain.picTempInv.visible Then DrawInventoryItem frmMain.picTempInv.left, frmMain.picTempInv.top
        If frmMain.picTempSpell.visible Then DrawDraggedSpell frmMain.picTempSpell.left, frmMain.picTempSpell.top
        If frmMain.picSpellDesc.visible Then DrawSpellDesc LastSpellDesc
        If frmMain.picItemDesc.visible Then DrawItemDesc LastItemDesc
        If frmMain.picHotbar.visible Then DrawHotbar
        If frmMain.picInventory.visible Then DrawInventory
        If frmMain.picItemDesc.visible Then DrawItemDesc LastItemDesc

        If frmMain.picSpells.visible Then DrawPlayerSpells
        If frmMain.picShop.visible Then DrawShop
        If frmMain.picTempBank.visible Then DrawBankItem frmMain.picTempBank.left, frmMain.picTempBank.top
        If frmMain.picBank.visible Then DrawBank
        If frmMain.picTrade.visible Then DrawTrade
    End If
    
    
    If frmEditor_Animation.visible Then
        EditorAnim_DrawAnim
    End If
    
    If frmEditor_Item.visible Then
        EditorItem_DrawItem
        EditorItem_DrawPaperdoll
    End If
    
    If frmEditor_Map.visible Then
        EditorMap_DrawTileset
        If frmEditor_Map.fraMapItem.visible Then EditorMap_DrawMapItem
        If frmEditor_Map.fraMapKey.visible Then EditorMap_DrawKey
    End If
    
    If frmEditor_NPC.visible Then
        EditorNpc_DrawSprite
    End If
    
    If frmEditor_Resource.visible Then
        EditorResource_DrawSprite
    End If
    
    If frmEditor_Spell.visible Then
        EditorSpell_DrawIcon
    End If
    
    If frmEditor_Events.visible Then
        EditorEvent_DrawGraphic
    End If
End Sub

Private Sub DrawFace()
Dim rec As RECT, rec_pos As RECT, faceNum As Long, srcRECT As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If NumFaces = 0 Then Exit Sub

    faceNum = GetPlayerSprite(MyIndex)
    
    If faceNum <= 0 Or faceNum > NumFaces Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    With rec
        .top = 0
        .Bottom = 100
        .left = 0
        .Right = 100
    End With

    With rec_pos
        .top = 0
        .Bottom = 100
        .left = 0
        .Right = 100
    End With

    RenderTextureByRects Tex_Face(faceNum), rec, rec_pos
    With srcRECT
        .x1 = 0
        .x2 = frmMain.picFace.Width
        .y1 = 0
        .y2 = frmMain.picFace.height
    End With
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRECT, srcRECT, frmMain.picFace.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawFace", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawEquipment()
Dim i As Long, itemNum As Long, ItemPic As Long
Dim rec As RECT, rec_pos As RECT, srcRECT As D3DRECT, destRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If numitems = 0 Then Exit Sub
    
    'frmMain.picCharacter.Cls
    For i = 1 To Equipment.Equipment_Count - 1
        itemNum = GetPlayerEquipment(MyIndex, i)

        If itemNum > 0 Then
            ItemPic = Item(itemNum).Pic

            With rec
                .top = 0
                .Bottom = 32
                .left = 32
                .Right = 64
            End With

            With rec_pos
                .top = EqTop
                .Bottom = .top + PIC_Y
                .left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .left + PIC_X
            End With
            Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
            Direct3D_Device.BeginScene
            RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
            Direct3D_Device.EndScene
            With srcRECT
                .x1 = rec_pos.left
                .x2 = rec_pos.Right
                .y1 = rec_pos.top
                .y2 = rec_pos.Bottom
            End With
            Direct3D_Device.Present srcRECT, srcRECT, frmMain.picCharacter.hwnd, ByVal (0)
        End If
    Next
    

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawEquipment", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawInventory()
Dim i As Long, x As Long, Y As Long, itemNum As Long, ItemPic As Long
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT, srcRECT As D3DRECT, destRect As D3DRECT
Dim colour As Long
Dim tmpItem As Long, amountModifier As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    ' reset gold label
    'frmMain.lblGold.Caption = "0g"
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    For i = 1 To MAX_INV
        itemNum = GetPlayerInvItemNum(MyIndex, i)

        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            
            amountModifier = 0
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For x = 1 To MAX_INV
                    tmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(x).num)
                    If TradeYourOffer(x).num = i Then
                        ' check if currency
                        If Not Item(tmpItem).Type = ITEM_TYPE_CURRENCY Then
                            ' normal item, exit out
                            GoTo NextLoop
                        Else
                            ' if amount = all currency, remove from inventory
                            If TradeYourOffer(x).Value = GetPlayerInvItemValue(MyIndex, i) Then
                                GoTo NextLoop
                            Else
                                ' not all, change modifier to show change in currency count
                                amountModifier = TradeYourOffer(x).Value
                            End If
                        End If
                    End If
                Next
            End If

            If ItemPic > 0 And ItemPic <= numitems Then
                If Tex_Item(ItemPic).Width <= 64 Then ' more than 1 frame is handled by anim sub

                    With rec
                        .top = 0
                        .Bottom = 32
                        .left = 32
                        .Right = 64
                    End With

                    With rec_pos
                        .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .Bottom = .top + PIC_Y
                        .left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .left + PIC_X
                    End With

                    RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        Y = rec_pos.top + 22
                        x = rec_pos.left - 4
                        
                        Amount = GetPlayerInvItemValue(MyIndex, i) - amountModifier
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If Amount < 1000000 Then
                            colour = White
                        ElseIf Amount > 1000000 And Amount < 10000000 Then
                            colour = Yellow
                        ElseIf Amount > 10000000 Then
                            colour = BrightGreen
                        End If
                        RenderText Font_Default, Format$(ConvertCurrency(str(Amount)), "#,###,###,###"), x, Y, colour, 0
                        ' Check if it's gold, and update the label
                        If GetPlayerInvItemNum(MyIndex, i) = 1 Then '1 = gold :P
                            frmMain.lblGold.Caption = Format$(Amount, "#,###,###,###") & "g"
                        End If
                    End If
                End If
            End If
        End If
NextLoop:
    Next
    
    'update animated items
    DrawAnimatedInvItems
    
    With srcRECT
        .x1 = 0
        .x2 = frmMain.picInventory.Width
        .y1 = 28
        .y2 = frmMain.picInventory.height + .y1
    End With
    
    With destRect
        .x1 = 0
        .x2 = frmMain.picInventory.Width
        .y1 = 32
        .y2 = frmMain.picInventory.height + .y1
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRECT, destRect, frmMain.picInventory.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawInventory", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub DrawHotbar()
Dim sRect As RECT, dRect As RECT, i As Long, num As String, n As Long, destRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 1 To MAX_HOTBAR
    
        With dRect
            .top = HotbarTop
            .left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
            .Bottom = .top + 32
            .Right = .left + 32
        End With
        
        With destRect
            .y1 = HotbarTop
            .x1 = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
            .y2 = .y1 + 32
            .x2 = .x1 + 32
        End With
        
        With sRect
            .top = 0
            .left = 32
            .Bottom = 32
            .Right = 64
        End With
        
        Select Case Hotbar(i).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(i).Slot).name) > 0 Then
                    If Item(Hotbar(i).Slot).Pic > 0 Then
                        If Item(Hotbar(i).Slot).Pic <= numitems Then
                            Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
                            Direct3D_Device.BeginScene
                            RenderTextureByRects Tex_Item(Item(Hotbar(i).Slot).Pic), sRect, dRect
                            Direct3D_Device.EndScene
                            Direct3D_Device.Present destRect, destRect, frmMain.picHotbar.hwnd, ByVal (0)
                        End If
                    End If
                End If
            Case 2 ' spell
                With sRect
                    .top = 0
                    .left = 0
                    .Bottom = 32
                    .Right = 32
                End With
                If Len(Spell(Hotbar(i).Slot).name) > 0 Then
                    If Spell(Hotbar(i).Slot).Icon > 0 Then
                        If Spell(Hotbar(i).Slot).Icon <= NumSpellIcons Then
                            ' check for cooldown
                            For n = 1 To MAX_PLAYER_SPELLS
                                If PlayerSpells(n) = Hotbar(i).Slot Then
                                    ' has spell
                                    If Not SpellCD(i) = 0 Then
                                        sRect.left = 32
                                        sRect.Right = 64
                                    End If
                                End If
                            Next
                            Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
                            Direct3D_Device.BeginScene
                            RenderTextureByRects Tex_SpellIcon(Spell(Hotbar(i).Slot).Icon), sRect, dRect
                            Direct3D_Device.EndScene
                            Direct3D_Device.Present destRect, destRect, frmMain.picHotbar.hwnd, ByVal (0)
                        End If
                    End If
                End If
        End Select
        
        ' render the letters
        num = "F" & str(i)
        RenderText Font_Default, num, dRect.left + 2, dRect.top + 16, White, 0
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawHotbar", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub DrawBank()
Dim i As Long, x As Long, Y As Long, itemNum As Long, srcRECT As D3DRECT, destRect As D3DRECT
Dim Amount As String
Dim sRect As RECT, dRect As RECT
Dim Sprite As Long, colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmMain.picBank.visible = True Then
        'frmMain.picBank.Cls
        
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
                
        For i = 1 To MAX_BANK
            itemNum = GetBankItemNum(i)
            If itemNum > 0 And itemNum <= MAX_ITEMS Then
            
                Sprite = Item(itemNum).Pic
                
                If Sprite <= 0 Or Sprite > numitems Then Exit Sub
            
                With sRect
                    .top = 0
                    .Bottom = .top + PIC_Y
                    .left = Tex_Item(Sprite).Width / 2
                    .Right = .left + PIC_X
                End With
                
                With dRect
                    .top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                    .Bottom = .top + PIC_Y
                    .left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                    .Right = .left + PIC_X
                End With
                
                RenderTextureByRects Tex_Item(Sprite), sRect, dRect

                ' If item is a stack - draw the amount you have
                If GetBankItemValue(i) > 1 Then
                    Y = dRect.top + 22
                    x = dRect.left - 4
                
                    Amount = CStr(GetBankItemValue(i))
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = BrightGreen
                    End If
                    RenderText Font_Default, ConvertCurrency(Amount), x, Y, colour
                End If
            End If
        Next
        
        With srcRECT
            .x1 = BankLeft
            .x2 = .x1 + 400
            .y1 = BankTop
            .y2 = .y1 + 310
        End With
                    
        With destRect
            .x1 = BankLeft
            .x2 = .x1 + 400
            .y1 = BankTop
            .y2 = 310 + .y1
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRECT, destRect, frmMain.picBank.hwnd, ByVal (0)
        'frmMain.picBank.Refresh
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBank", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawBankItem(ByVal x As Long, ByVal Y As Long)
Dim sRect As RECT, dRect As RECT, srcRECT As D3DRECT, destRect As D3DRECT
Dim itemNum As Long
Dim Sprite As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = GetBankItemNum(DragBankSlotNum)
    Sprite = Item(GetBankItemNum(DragBankSlotNum)).Pic
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
    Direct3D_Device.BeginScene
    
    If itemNum > 0 Then
        If itemNum <= MAX_ITEMS Then
            With sRect
                .top = 0
                .Bottom = .top + PIC_Y
                .left = Tex_Item(Sprite).Width / 2
                .Right = .left + PIC_X
            End With
        End If
    End If
    
    With dRect
        .top = 2
        .Bottom = .top + PIC_Y
        .left = 2
        .Right = .left + PIC_X
    End With

    RenderTextureByRects Tex_Item(Sprite), sRect, dRect
    
    With frmMain.picTempBank
        .top = Y
        .left = x
        .visible = True
        .ZOrder (0)
    End With
    
    With srcRECT
        .x1 = 0
        .x2 = 32
        .y1 = 0
        .y2 = 32
    End With
    With destRect
        .x1 = 2
        .y1 = 2
        .y2 = .y1 + 32
        .x2 = .x1 + 32
    End With
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRECT, destRect, frmMain.picTempBank.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBankItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub DrawInventoryItem(ByVal x As Long, ByVal Y As Long)
Dim rec As RECT, rec_pos As RECT, srcRECT As D3DRECT, destRect As D3DRECT
Dim itemNum As Long, ItemPic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    itemNum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)

    If itemNum > 0 And itemNum <= MAX_ITEMS Then
        ItemPic = Item(itemNum).Pic
        
        If ItemPic = 0 Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
        Direct3D_Device.BeginScene
        
        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .left = Tex_Item(ItemPic).Width / 2
            .Right = .left + PIC_X
        End With

        With rec_pos
            .top = 2
            .Bottom = .top + PIC_Y
            .left = 2
            .Right = .left + PIC_X
        End With

        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

        With frmMain.picTempInv
            .top = Y
            .left = x
            .visible = True
            .ZOrder (0)
        End With
        With srcRECT
            .x1 = 0
            .x2 = 32
            .y1 = 0
            .y2 = 32
        End With
        With destRect
            .x1 = 2
            .y1 = 2
            .y2 = .y1 + 32
            .x2 = .x1 + 32
        End With
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRECT, destRect, frmMain.picTempInv.hwnd, ByVal (0)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawInventoryItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawDraggedSpell(ByVal x As Long, ByVal Y As Long)
Dim rec As RECT, rec_pos As RECT, srcRECT As D3DRECT, destRect As D3DRECT
Dim spellnum As Long, spellpic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    spellnum = PlayerSpells(DragSpell)

    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        spellpic = Spell(spellnum).Icon
        
        If spellpic = 0 Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .left = 0
            .Right = .left + PIC_X
        End With

        With rec_pos
            .top = 2
            .Bottom = .top + PIC_Y
            .left = 2
            .Right = .left + PIC_X
        End With

        RenderTextureByRects Tex_SpellIcon(spellpic), rec, rec_pos

        With frmMain.picTempSpell
            .top = Y
            .left = x
            .visible = True
            .ZOrder (0)
        End With
        
        With srcRECT
            .x1 = 0
            .x2 = 32
            .y1 = 0
            .y2 = 32
        End With
        With destRect
            .x1 = 2
            .y1 = 2
            .y2 = .y1 + 32
            .x2 = .x1 + 32
        End With
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRECT, destRect, frmMain.picTempSpell.hwnd, ByVal (0)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawInventoryItem", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawItemDesc(ByVal itemNum As Long)
Dim rec As RECT, rec_pos As RECT, srcRECT As D3DRECT, destRect As D3DRECT
Dim ItemPic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'frmMain.picItemDescPic.Cls
    
    If itemNum > 0 And itemNum <= MAX_ITEMS Then
        ItemPic = Item(itemNum).Pic

        If ItemPic = 0 Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .left = Tex_Item(ItemPic).Width / 2
            .Right = .left + PIC_X
        End With

        With rec_pos
            .top = 0
            .Bottom = 64
            .left = 0
            .Right = 64
        End With
        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

        With destRect
            .x1 = 0
            .y1 = 0
            .y2 = 64
            .x2 = 64
        End With
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present destRect, destRect, frmMain.picItemDescPic.hwnd, ByVal (0)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawItemDesc", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawSpellDesc(ByVal spellnum As Long)
Dim rec As RECT, rec_pos As RECT, srcRECT As D3DRECT, destRect As D3DRECT
Dim spellpic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    'frmMain.picSpellDescPic.Cls

    If spellnum > 0 And spellnum <= MAX_SPELLS Then
        spellpic = Spell(spellnum).Icon

        If spellpic <= 0 Or spellpic > NumSpellIcons Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene

        With rec
            .top = 0
            .Bottom = .top + PIC_Y
            .left = 0
            .Right = .left + PIC_X
        End With

        With rec_pos
            .top = 0
            .Bottom = 64
            .left = 0
            .Right = 64
        End With
        RenderTextureByRects Tex_SpellIcon(spellpic), rec, rec_pos

        With destRect
            .x1 = 0
            .y1 = 0
            .y2 = 64
            .x2 = 64
        End With
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present destRect, destRect, frmMain.picSpellDescPic.hwnd, ByVal (0)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawSpellDesc", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub DrawTrade()
Dim i As Long, x As Long, Y As Long, itemNum As Long, ItemPic As Long, srcRECT As D3DRECT, destRect As D3DRECT
Dim Amount As Long
Dim rec As RECT, rec_pos As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    For i = 1 To MAX_INV
        ' Draw your own offer
        itemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)

        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic

            If ItemPic > 0 And ItemPic <= numitems Then
                With rec
                    .top = 0
                    .Bottom = 32
                    .left = 32
                    .Right = 64
                End With

                With rec_pos
                    .top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .top + PIC_Y
                    .left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .left + PIC_X
                End With

                RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).Value > 1 Then
                    Y = rec_pos.top + 22
                    x = rec_pos.left - 4
                    
                    Amount = TradeYourOffer(i).Value
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        colour = White
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        colour = Yellow
                    ElseIf Amount > 10000000 Then
                        colour = BrightGreen
                    End If
                    RenderText Font_Default, ConvertCurrency(str(Amount)), x, Y, colour, 0
                End If
            End If
        End If
    Next
    
    With srcRECT
        .x1 = 0
        .x2 = .x1 + 193
        .y1 = 0
        .y2 = .y1 + 246
    End With
                    
    With destRect
        .x1 = 0
        .x2 = .x1 + 193
        .y1 = 0
        .y2 = 246 + .y1
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRECT, destRect, frmMain.picYourTrade.hwnd, ByVal (0)
    
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    For i = 1 To MAX_INV
        ' Draw their offer
        itemNum = TradeTheirOffer(i).num

        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic

            If ItemPic > 0 And ItemPic <= numitems Then
                With rec
                    .top = 0
                    .Bottom = 32
                    .left = 32
                    .Right = 64
                End With

                With rec_pos
                    .top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .Bottom = .top + PIC_Y
                    .left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .left + PIC_X
                End With
                
                RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).Value > 1 Then
                    Y = rec_pos.top + 22
                    x = rec_pos.left - 4
                    
                    Amount = TradeTheirOffer(i).Value
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        colour = White
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        colour = Yellow
                    ElseIf Amount > 10000000 Then
                        colour = BrightGreen
                    End If
                    RenderText Font_Default, ConvertCurrency(str(Amount)), x, Y, colour, 0
                End If
            End If
        End If
    Next
    
    With srcRECT
        .x1 = 0
        .x2 = .x1 + 193
        .y1 = 0
        .y2 = .y1 + 246
    End With
                    
    With destRect
        .x1 = 0
        .x2 = .x1 + 193
        .y1 = 0
        .y2 = 246 + .y1
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRECT, destRect, frmMain.picTheirTrade.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTrade", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawPlayerSpells()
Dim i As Long, x As Long, Y As Long, spellnum As Long, spellicon As Long, srcRECT As D3DRECT, destRect As D3DRECT
Dim Amount As String
Dim rec As RECT, rec_pos As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    'frmMain.picSpells.Cls
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i)

        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellicon = Spell(spellnum).Icon

            If spellicon > 0 And spellicon <= NumSpellIcons Then
            
                With rec
                    .top = 0
                    .Bottom = 32
                    .left = 0
                    .Right = 32
                End With
                
                If Not SpellCD(i) = 0 Then
                    rec.left = 32
                    rec.Right = 64
                End If

                With rec_pos
                    .top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                    .Bottom = .top + PIC_Y
                    .left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                    .Right = .left + PIC_X
                End With

                RenderTextureByRects Tex_SpellIcon(spellicon), rec, rec_pos
            End If
        End If
    Next
    
    With srcRECT
        .x1 = 0
        .x2 = frmMain.picSpells.Width
        .y1 = 28
        .y2 = frmMain.picSpells.height + .y1
    End With
    
    With destRect
        .x1 = 0
        .x2 = frmMain.picSpells.Width
        .y1 = 32
        .y2 = frmMain.picSpells.height + .y1
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRECT, destRect, frmMain.picSpells.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerSpells", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawShop()
Dim i As Long, x As Long, Y As Long, itemNum As Long, ItemPic As Long, srcRECT As D3DRECT, destRect As D3DRECT
Dim Amount As String
Dim rec As RECT, rec_pos As RECT
Dim colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not InGame Then Exit Sub
    
    'frmMain.picShopItems.Cls
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    For i = 1 To MAX_TRADES
        itemNum = Shop(InShop).TradeItem(i).Item 'GetPlayerInvItemNum(MyIndex, i)
        If itemNum > 0 And itemNum <= MAX_ITEMS Then
            ItemPic = Item(itemNum).Pic
            If ItemPic > 0 And ItemPic <= numitems Then
            
                With rec
                    .top = 0
                    .Bottom = 32
                    .left = 32
                    .Right = 64
                End With
                
                With rec_pos
                    .top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                    .Bottom = .top + PIC_Y
                    .left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                    .Right = .left + PIC_X
                End With
                
                RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
                
                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                    Y = rec_pos.top + 22
                    x = rec_pos.left - 4
                    Amount = CStr(Shop(InShop).TradeItem(i).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = Green
                    End If
                    RenderText Font_Default, ConvertCurrency(Amount), x, Y, colour, 0
                End If
            End If
        End If
    Next
    
    With srcRECT
        .x1 = ShopLeft
        .x2 = .x1 + 192
        .y1 = ShopTop
        .y2 = .y1 + 211
    End With
                
    With destRect
        .x1 = ShopLeft
        .x2 = .x1 + 192
        .y1 = ShopTop
        .y2 = 211 + .y1
    End With
                
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRECT, destRect, frmMain.picShopItems.hwnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawShop", "modGraphics", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
