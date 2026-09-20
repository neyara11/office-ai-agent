Imports System.Drawing

Namespace Design

    Public NotInheritable Partial Class SlideComponentLibrary

        Private Shared Sub BuildEditorialListContent(plan As SlideRenderPlan,
                                                     items As List(Of DesignItem),
                                                     t As DesignTokens,
                                                     w As Single,
                                                     h As Single)
            Dim left As Single = 62, top As Single = 166, bottom As Single = 58
            Dim contentHeight = h - top - bottom
            Dim rowHeight = contentHeight / Math.Max(1, items.Count)
            AddRect(plan, "editorial-accent", left, top, 6, contentHeight, t.Primary, collision:=False)
            For index = 0 To items.Count - 1
                Dim y = top + index * rowHeight
                Dim fontSize = If(items.Count <= 2, t.TitleSize - 2, t.BodySize + 1)
                AddText(plan, $"editorial-index-{index + 1}", (index + 1).ToString("00"), left + 24, y + 15, 42, 24, t.CaptionSize, t.Primary, True)
                AddText(plan, $"editorial-item-{index + 1}-title", items(index).Title,
                        left + 84, y + 10, w - left - 140, rowHeight - 20,
                        fontSize, t.TextPrimary, index = 0)
                If index < items.Count - 1 Then
                    AddLine(plan, $"editorial-divider-{index + 1}", left + 84, y + rowHeight - 2,
                            w - 62, y + rowHeight - 2, t.Divider, 0.8F)
                End If
            Next
        End Sub

        Private Shared Sub BuildImageContent(plan As SlideRenderPlan,
                                             imagePath As String,
                                             items As List(Of DesignItem),
                                             t As DesignTokens,
                                             w As Single,
                                             h As Single)
            Dim left As Single = 58, top As Single = 163, bottom As Single = 58, gap As Single = 22
            Dim imageWidth = (w - left * 2) * 0.46F
            Dim contentHeight = h - top - bottom
            AddRect(plan, "image-frame", left, top, imageWidth, contentHeight, t.Surface, t.Divider, collision:=False)
            AddImage(plan, "content-image", imagePath, left + 8, top + 8, imageWidth - 16, contentHeight - 16)

            Dim rightX = left + imageWidth + gap
            Dim rightWidth = w - left - rightX
            Dim visibleItems = items.Take(4).ToList()
            Dim rowGap As Single = 12
            Dim rowHeight = (contentHeight - rowGap * Math.Max(0, visibleItems.Count - 1)) / Math.Max(1, visibleItems.Count)
            For index = 0 To visibleItems.Count - 1
                Dim y = top + index * (rowHeight + rowGap)
                AddRect(plan, $"image-point-{index + 1}", rightX, y, rightWidth, rowHeight, t.Surface, t.Divider)
                AddText(plan, $"image-point-index-{index + 1}", (index + 1).ToString("00"), rightX + 16, y + 14, 34, 20, t.CaptionSize, t.Primary, True)
                If visibleItems.Count >= 4 Then
                    AddText(plan, $"image-point-compact-{index + 1}", ComposeItemText(visibleItems(index)),
                            rightX + 60, y + 12, rightWidth - 76, rowHeight - 24,
                            t.CaptionSize + 1, t.TextPrimary, index = 0)
                Else
                    AddText(plan, $"image-point-title-{index + 1}", visibleItems(index).Title, rightX + 60, y + 12, rightWidth - 76, 30, t.BodySize, t.TextPrimary, True)
                    AddText(plan, $"image-point-body-{index + 1}", visibleItems(index).Body, rightX + 60, y + 45, rightWidth - 76, rowHeight - 54, t.CaptionSize, t.TextSecondary, False)
                End If
            Next
        End Sub

        Private Shared Sub BuildFeatureContent(plan As SlideRenderPlan,
                                               items As List(Of DesignItem),
                                               t As DesignTokens,
                                               w As Single,
                                               h As Single)
            Dim left As Single = 58, top As Single = 163, bottom As Single = 58, gap As Single = 20
            Dim featureWidth = (w - left * 2) * 0.43F
            Dim featureHeight = h - top - bottom
            Dim feature = items.FirstOrDefault(Function(item) item.Emphasis)
            If feature Is Nothing Then feature = items(0)
            AddRect(plan, "feature-card", left, top, featureWidth, featureHeight, t.SurfaceAlt, t.Primary, shadow:=True)
            AddText(plan, "feature-title", feature.Title, left + 24, top + 32, featureWidth - 48, 78, t.TitleSize - 3, t.TextPrimary, True)
            AddText(plan, "feature-body", feature.Body, left + 24, top + 128, featureWidth - 48, featureHeight - 156, t.BodySize, t.TextSecondary, False)

            Dim remaining = items.Where(Function(item) Not Object.ReferenceEquals(item, feature)).Take(4).ToList()
            Dim rightX = left + featureWidth + gap
            Dim rightWidth = w - left - rightX
            Dim rowGap As Single = 12
            Dim rowHeight = (featureHeight - rowGap * Math.Max(0, remaining.Count - 1)) / Math.Max(1, remaining.Count)
            For index = 0 To remaining.Count - 1
                Dim y = top + index * (rowHeight + rowGap)
                AddRect(plan, $"feature-side-{index + 1}", rightX, y, rightWidth, rowHeight, t.Surface, t.Divider)
                AddText(plan, $"feature-side-index-{index + 1}", (index + 2).ToString("00"), rightX + 18, y + 16, 36, 20, t.CaptionSize, t.Primary, True)
                If remaining.Count >= 4 Then
                    AddText(plan, $"feature-side-compact-{index + 1}", ComposeItemText(remaining(index)),
                            rightX + 66, y + 12, rightWidth - 84, rowHeight - 24,
                            t.CaptionSize + 1, t.TextPrimary, index = 0)
                Else
                    AddText(plan, $"feature-side-title-{index + 1}", remaining(index).Title, rightX + 66, y + 13, rightWidth - 84, 30, t.BodySize, t.TextPrimary, True)
                    AddText(plan, $"feature-side-body-{index + 1}", remaining(index).Body, rightX + 66, y + 47, rightWidth - 84, rowHeight - 58, t.CaptionSize, t.TextSecondary, False)
                End If
            Next
        End Sub

        Private Shared Sub BuildTwoColumn(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single, h As Single)
            AddHeader(plan, spec, t, w)
            Dim items = spec.Items.Take(2).ToList()
            Dim margin As Single = 58, gap As Single = 22, top As Single = 160, cardH = h - top - 58
            Dim cardW = (w - margin * 2 - gap) / 2
            For index = 0 To 1
                AddCard(plan, $"column-{index + 1}", items(index), margin + index * (cardW + gap), top, cardW, cardH, t, index)
            Next
        End Sub

        Private Shared Sub BuildComparison(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single, h As Single)
            AddHeader(plan, spec, t, w)
            If spec.Items.Count >= 3 AndAlso spec.Items.All(Function(item) item.Features IsNot Nothing AndAlso item.Features.Count >= 2) Then
                BuildComparisonTable(plan, spec.Items.Take(5).ToList(), spec.ColumnHeaders, t, w, h)
                Return
            End If
            Dim items = spec.Items.Take(2).ToList()
            Dim margin As Single = 58, gap As Single = 20, top As Single = 165, cardH = h - top - 58
            Dim totalWidth = w - margin * 2 - gap
            Dim emphasizedIndex = items.FindIndex(Function(item) item.Emphasis)
            Dim firstWidth = If(emphasizedIndex = 0, totalWidth * 0.56F,
                                If(emphasizedIndex = 1, totalWidth * 0.44F, totalWidth / 2.0F))
            Dim secondWidth = totalWidth - firstWidth
            For index = 0 To 1
                Dim cardWidth = If(index = 0, firstWidth, secondWidth)
                Dim x = If(index = 0, margin, margin + firstWidth + gap)
                Dim emphasized = index = emphasizedIndex
                Dim accent = If(emphasized, t.Primary, t.TextSecondary)
                AddRect(plan, $"compare-card-{index + 1}", x, top, cardWidth, cardH,
                        If(emphasized, t.SurfaceAlt, t.Surface), If(emphasized, t.Primary, t.Divider),
                        shadow:=emphasized)
                AddRect(plan, $"compare-bar-{index + 1}", x, top, cardWidth, 8, accent, collision:=False)
                AddText(plan, $"compare-title-{index + 1}", items(index).Title, x + 24, top + 30, cardWidth - 48, 42, t.TitleSize - 5, t.TextPrimary, True)
                AddText(plan, $"compare-body-{index + 1}", items(index).Body, x + 24, top + 92, cardWidth - 48, cardH - 118, t.BodySize, t.TextSecondary, False)
            Next
        End Sub

        Private Shared Sub BuildComparisonTable(plan As SlideRenderPlan,
                                                 items As List(Of DesignItem),
                                                 headers As List(Of String),
                                                 t As DesignTokens,
                                                w As Single,
                                                h As Single)
            Dim left As Single = 58, top As Single = 163, totalWidth = w - 116
            Dim labelWidth = totalWidth * 0.28F
            Dim valueWidth = (totalWidth - labelWidth) / 2
            Dim headerHeight As Single = 48
            Dim rowHeight = (h - top - 62 - headerHeight) / Math.Max(1, items.Count)
            AddRect(plan, "compare-table-header", left, top, totalWidth, headerHeight, t.SurfaceAlt, t.Divider)
            Dim dimensionHeader = headers(0)
            Dim leftHeader = headers(1)
            Dim rightHeader = headers(2)
            AddText(plan, "compare-dimension-header", dimensionHeader, left + 18, top + 14, labelWidth - 30, 24, t.CaptionSize, t.TextSecondary, True)
            AddText(plan, "compare-left-header", leftHeader, left + labelWidth + 16, top + 12, valueWidth - 28, 28, t.BodySize, t.TextPrimary, True, "center")
            AddText(plan, "compare-right-header", rightHeader, left + labelWidth + valueWidth + 16, top + 12, valueWidth - 28, 28, t.BodySize, t.Primary, True, "center")
            For index = 0 To items.Count - 1
                Dim y = top + headerHeight + index * rowHeight
                AddRect(plan, $"compare-row-{index + 1}", left, y, totalWidth, rowHeight - 4,
                        If(index Mod 2 = 0, t.Surface, t.Background), t.Divider)
                AddText(plan, $"compare-label-{index + 1}", items(index).Title, left + 18, y + 13, labelWidth - 30, rowHeight - 22, t.CaptionSize + 1, t.TextPrimary, True)
                Dim leftValue = If(items(index).Features.Count > 0, items(index).Features(0), items(index).Body)
                Dim rightValue = If(items(index).Features.Count > 1, items(index).Features(1), "")
                AddText(plan, $"compare-left-{index + 1}", leftValue, left + labelWidth + 16, y + 12, valueWidth - 28, rowHeight - 20, t.CaptionSize + 1, t.TextSecondary, False, "center")
                AddText(plan, $"compare-right-{index + 1}", rightValue, left + labelWidth + valueWidth + 16, y + 12, valueWidth - 28, rowHeight - 20, t.CaptionSize + 1, t.TextPrimary, True, "center")
            Next
        End Sub

        Private Shared Sub BuildKpi(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single, h As Single)
            AddHeader(plan, spec, t, w)
            Dim metrics = spec.Metrics.Take(4).ToList()
            If metrics.Count = 0 Then
                For Each item In spec.Items.Take(4)
                    metrics.Add(New DesignMetric With {.Value = item.Value, .Label = item.Title, .Description = item.Body})
                Next
            End If
            If String.Equals(spec.LayoutVariant, "hero-left", StringComparison.OrdinalIgnoreCase) AndAlso metrics.Count >= 3 Then
                BuildKpiHeroLeft(plan, spec, metrics, t, w, h)
                Return
            End If
            Dim margin As Single = 58, gap As Single = 14, top As Single = 178, cardH As Single = 198
            Dim cardW = (w - margin * 2 - gap * Math.Max(0, metrics.Count - 1)) / Math.Max(1, metrics.Count)
            For index = 0 To metrics.Count - 1
                Dim x = margin + index * (cardW + gap)
                AddRect(plan, $"metric-card-{index + 1}", x, top, cardW, cardH, t.Surface, t.Divider)
                AddText(plan, $"metric-value-{index + 1}", metrics(index).Value, x + 20, top + 28, cardW - 40, 58, t.DisplaySize - 2, If(index = 0, t.Primary, t.TextPrimary), True)
                AddText(plan, $"metric-label-{index + 1}", metrics(index).Label, x + 20, top + 94, cardW - 40, 30, t.BodySize, t.TextPrimary, True)
                AddText(plan, $"metric-desc-{index + 1}", FirstNonEmpty(metrics(index).Delta, metrics(index).Description), x + 20, top + 132, cardW - 40, 38, t.CaptionSize, t.TextSecondary, False)
                AddText(plan, $"metric-source-{index + 1}", metrics(index).Source, x + 20, top + 178, cardW - 40, 14, 9.5F, t.TextSecondary, False)
            Next
            AddText(plan, "kpi-insight", spec.Body, 60, 395, w - 120, 66, t.BodySize + 1, t.TextSecondary, False, "center")
        End Sub

        Private Shared Sub BuildKpiHeroLeft(plan As SlideRenderPlan,
                                            spec As SlideDesignSpec,
                                            metrics As List(Of DesignMetric),
                                            t As DesignTokens,
                                            w As Single,
                                            h As Single)
            Dim margin As Single = 58, top As Single = 170, bottom As Single = 58, gap As Single = 18
            Dim contentHeight = h - top - bottom
            Dim heroWidth = (w - margin * 2 - gap) * 0.43F
            AddRect(plan, "metric-hero", margin, top, heroWidth, contentHeight, t.SurfaceAlt, t.Primary, shadow:=True)
            AddText(plan, "metric-value-1", metrics(0).Value, margin + 28, top + 30, heroWidth - 56, 78,
                    t.DisplaySize + 12, t.Primary, True)
            AddText(plan, "metric-label-1", metrics(0).Label, margin + 28, top + 120, heroWidth - 56, 42,
                    t.TitleSize - 3, t.TextPrimary, True)
            AddText(plan, "metric-desc-1", FirstNonEmpty(metrics(0).Delta, metrics(0).Description),
                    margin + 28, top + 176, heroWidth - 56, 52, t.BodySize, t.TextSecondary, False)
            AddText(plan, "kpi-insight", spec.Body, margin + 28, top + contentHeight - 76,
                    heroWidth - 56, 38, t.CaptionSize + 1, t.TextSecondary, False)
            AddText(plan, "metric-source-1", metrics(0).Source, margin + 28, top + contentHeight - 28,
                    heroWidth - 56, 14, 9.5F, t.TextSecondary, False)

            Dim remaining = metrics.Skip(1).Take(3).ToList()
            Dim rightX = margin + heroWidth + gap
            Dim rightWidth = w - margin - rightX
            Dim rowGap As Single = 12
            Dim rowHeight = (contentHeight - rowGap * Math.Max(0, remaining.Count - 1)) / Math.Max(1, remaining.Count)
            For index = 0 To remaining.Count - 1
                Dim metric = remaining(index)
                Dim y = top + index * (rowHeight + rowGap)
                AddRect(plan, $"metric-side-{index + 2}", rightX, y, rightWidth, rowHeight, t.Surface, t.Divider)
                AddText(plan, $"metric-value-{index + 2}", metric.Value, rightX + 22, y + 18,
                        rightWidth * 0.3F, 48, t.DisplaySize - 6, t.TextPrimary, True)
                AddText(plan, $"metric-label-{index + 2}", metric.Label, rightX + rightWidth * 0.35F, y + 18,
                        rightWidth * 0.58F, 30, t.BodySize + 1, t.TextPrimary, True)
                AddText(plan, $"metric-desc-{index + 2}", FirstNonEmpty(metric.Delta, metric.Description),
                        rightX + rightWidth * 0.35F, y + 52, rightWidth * 0.58F, rowHeight - 78,
                        t.CaptionSize + 1, t.TextSecondary, False)
                AddText(plan, $"metric-source-{index + 2}", metric.Source,
                        rightX + rightWidth * 0.35F, y + rowHeight - 20, rightWidth * 0.58F, 12,
                        9.5F, t.TextSecondary, False)
            Next
        End Sub

        Private Shared Sub BuildProcess(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single, h As Single)
            AddHeader(plan, spec, t, w)
            Dim items = spec.Items.Take(6).ToList()
            If String.Equals(spec.LayoutVariant, "vertical", StringComparison.OrdinalIgnoreCase) Then
                BuildVerticalProcess(plan, items, t, w, h)
                Return
            End If
            Dim margin As Single = 66, top As Single = 245
            Dim available = w - margin * 2
            Dim stepGap = available / Math.Max(1, items.Count - 1)
            AddLine(plan, "process-line", margin, top, w - margin, top, t.Divider, 3)
            For index = 0 To items.Count - 1
                Dim centerX = If(items.Count = 1, w / 2, margin + index * stepGap)
                AddCircle(plan, $"step-dot-{index + 1}", centerX - 23, top - 23, 46, If(index = 0, t.Primary, t.SurfaceAlt), 0)
                AddText(plan, $"step-number-{index + 1}", (index + 1).ToString("00"), centerX - 20, top - 13, 40, 22, t.CaptionSize, If(index = 0, t.Background, t.TextPrimary), True, "center")
                AddText(plan, $"step-title-{index + 1}", items(index).Title, centerX - 68, top + 47, 136, 38, t.BodySize, t.TextPrimary, True, "center")
                AddText(plan, $"step-body-{index + 1}", items(index).Body, centerX - 72, top + 88, 144, 58, t.CaptionSize, t.TextSecondary, False, "center")
            Next
        End Sub

        Private Shared Sub BuildVerticalProcess(plan As SlideRenderPlan,
                                                items As List(Of DesignItem),
                                                t As DesignTokens,
                                                w As Single,
                                                h As Single)
            Dim top As Single = 164, bottom As Single = 58, left As Single = 72
            Dim contentHeight = h - top - bottom
            Dim rowHeight = contentHeight / Math.Max(1, items.Count)
            AddLine(plan, "process-vertical-line", left + 18, top + 18,
                    left + 18, top + contentHeight - 18, t.Divider, 2.5F)
            For index = 0 To items.Count - 1
                Dim y = top + index * rowHeight
                AddCircle(plan, $"step-dot-{index + 1}", left, y + rowHeight / 2 - 18, 36,
                          If(index = 0, t.Primary, t.SurfaceAlt), 0)
                AddText(plan, $"step-number-{index + 1}", (index + 1).ToString("00"),
                        left + 2, y + rowHeight / 2 - 9, 32, 18, t.CaptionSize,
                        If(index = 0, t.Background, t.TextPrimary), True, "center")
                AddText(plan, $"step-title-{index + 1}", items(index).Title,
                        left + 68, y + 12, w * 0.28F, rowHeight - 24,
                        t.BodySize + 1, t.TextPrimary, True)
                AddText(plan, $"step-body-{index + 1}", items(index).Body,
                        left + w * 0.35F, y + 12, w - left - w * 0.35F - 62, rowHeight - 24,
                        t.CaptionSize + 1, t.TextSecondary, False)
                If index < items.Count - 1 Then
                    AddLine(plan, $"step-divider-{index + 1}", left + 68, y + rowHeight - 1,
                            w - 58, y + rowHeight - 1, t.Divider, 0.7F)
                End If
            Next
        End Sub

        Private Shared Sub BuildArchitecture(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single, h As Single)
            AddHeader(plan, spec, t, w)
            If String.Equals(spec.LayoutVariant, "hub-spoke", StringComparison.OrdinalIgnoreCase) Then
                BuildHubSpokeArchitecture(plan, spec.Items.Take(5).ToList(), t, w, h)
                Return
            End If
            Dim items = spec.Items.Take(5).ToList()
            Dim margin As Single = 105, top As Single = 155, gap As Single = 13
            Dim layerH = (h - top - 60 - gap * (items.Count - 1)) / Math.Max(1, items.Count)
            For index = 0 To items.Count - 1
                Dim inset = index * 16
                Dim x = margin + inset
                Dim width = w - 2 * x
                Dim y = top + index * (layerH + gap)
                AddRect(plan, $"layer-{index + 1}", x, y, width, layerH, If(index Mod 2 = 0, t.Surface, t.SurfaceAlt), t.Divider)
                AddText(plan, $"layer-index-{index + 1}", (index + 1).ToString("00"), x + 18, y + 17, 42, 25, t.CaptionSize, t.Primary, True)
                AddText(plan, $"layer-title-{index + 1}", items(index).Title, x + 68, y + 12, width * 0.3F, 35, t.BodySize + 1, t.TextPrimary, True)
                AddText(plan, $"layer-body-{index + 1}", items(index).Body, x + width * 0.42F, y + 13, width * 0.53F, layerH - 20, t.CaptionSize + 1, t.TextSecondary, False)
            Next
        End Sub

        Private Shared Sub BuildHubSpokeArchitecture(plan As SlideRenderPlan,
                                                     items As List(Of DesignItem),
                                                     t As DesignTokens,
                                                     w As Single,
                                                     h As Single)
            Dim top As Single = 166, bottom As Single = 58, margin As Single = 58
            Dim contentHeight = h - top - bottom
            Dim coreWidth As Single = 240, coreHeight As Single = 140
            Dim coreX = (w - coreWidth) / 2.0F
            Dim coreY = top + (contentHeight - coreHeight) / 2.0F
            Dim centerY = coreY + coreHeight / 2.0F
            Dim spokeWidth As Single = 220
            Dim spokeHeight As Single = If(items.Count <= 3, 120, 112)
            Dim leftX = margin
            Dim rightX = w - margin - spokeWidth
            Dim topY = top
            Dim bottomY = top + contentHeight - spokeHeight
            Dim middleY = top + (contentHeight - spokeHeight) / 2.0F
            Dim spokeBounds As New List(Of SceneRect)()

            Select Case items.Count - 1
                Case 2
                    spokeBounds.Add(New SceneRect(leftX, middleY, spokeWidth, spokeHeight))
                    spokeBounds.Add(New SceneRect(rightX, middleY, spokeWidth, spokeHeight))
                Case 3
                    spokeBounds.Add(New SceneRect(leftX, middleY, spokeWidth, spokeHeight))
                    spokeBounds.Add(New SceneRect(rightX, topY, spokeWidth, spokeHeight))
                    spokeBounds.Add(New SceneRect(rightX, bottomY, spokeWidth, spokeHeight))
                Case Else
                    spokeBounds.Add(New SceneRect(leftX, topY, spokeWidth, spokeHeight))
                    spokeBounds.Add(New SceneRect(leftX, bottomY, spokeWidth, spokeHeight))
                    spokeBounds.Add(New SceneRect(rightX, topY, spokeWidth, spokeHeight))
                    spokeBounds.Add(New SceneRect(rightX, bottomY, spokeWidth, spokeHeight))
            End Select

            Dim leftSpokes = spokeBounds.Where(Function(bounds) bounds.X < coreX).ToList()
            Dim rightSpokes = spokeBounds.Where(Function(bounds) bounds.X > coreX).ToList()
            If leftSpokes.Count > 0 Then
                Dim branchX = leftX + spokeWidth
                AddLine(plan, "architecture-connector-left", branchX, centerY, coreX, centerY, t.Divider, 1.8F)
                For index = 0 To leftSpokes.Count - 1
                    Dim spokeCenterY = leftSpokes(index).Y + leftSpokes(index).Height / 2.0F
                    If Math.Abs(spokeCenterY - centerY) > 0.5F Then
                        AddLine(plan, $"architecture-branch-left-{index + 1}", branchX,
                                Math.Min(spokeCenterY, centerY), branchX,
                                Math.Max(spokeCenterY, centerY), t.Divider, 1.4F)
                    End If
                Next
            End If
            If rightSpokes.Count > 0 Then
                Dim branchX = rightX
                AddLine(plan, "architecture-connector-right", coreX + coreWidth, centerY, branchX, centerY, t.Divider, 1.8F)
                For index = 0 To rightSpokes.Count - 1
                    Dim spokeCenterY = rightSpokes(index).Y + rightSpokes(index).Height / 2.0F
                    If Math.Abs(spokeCenterY - centerY) > 0.5F Then
                        AddLine(plan, $"architecture-branch-right-{index + 1}", branchX,
                                Math.Min(spokeCenterY, centerY), branchX,
                                Math.Max(spokeCenterY, centerY), t.Divider, 1.4F)
                    End If
                Next
            End If

            AddRect(plan, "architecture-core", coreX, coreY, coreWidth, coreHeight,
                    t.SurfaceAlt, t.Primary, shadow:=True)
            AddText(plan, "architecture-core-title", items(0).Title,
                    coreX + 24, coreY + 20, coreWidth - 48, 58,
                    t.TitleSize - 5, t.TextPrimary, True, "center")
            AddText(plan, "architecture-core-body", items(0).Body,
                    coreX + 24, coreY + 88, coreWidth - 48, 32,
                    t.CaptionSize + 1, t.TextSecondary, False, "center")

            For index = 0 To spokeBounds.Count - 1
                Dim bounds = spokeBounds(index)
                Dim item = items(index + 1)
                AddRect(plan, $"architecture-spoke-{index + 1}", bounds.X, bounds.Y,
                        bounds.Width, bounds.Height, t.Surface, t.Divider)
                AddText(plan, $"architecture-spoke-{index + 1}-index", (index + 1).ToString("00"),
                        bounds.X + 16, bounds.Y + 14, 34, 20, t.CaptionSize, t.Primary, True)
                AddText(plan, $"architecture-spoke-{index + 1}-title", item.Title,
                        bounds.X + 58, bounds.Y + 12, bounds.Width - 74, 34,
                        t.BodySize, t.TextPrimary, True)
                AddText(plan, $"architecture-spoke-{index + 1}-body", item.Body,
                        bounds.X + 58, bounds.Y + 52, bounds.Width - 74, bounds.Height - 66,
                        t.CaptionSize, t.TextSecondary, False)
            Next
        End Sub

        Private Shared Sub BuildMatrix(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single, h As Single)
            AddHeader(plan, spec, t, w)
            Dim items = spec.Items.Take(4).ToList()
            Dim left As Single = 118, top As Single = 168, gap As Single = 12
            Dim cellW = (w - left - 76 - gap) / 2, cellH = (h - top - 58 - gap) / 2
            For index = 0 To 3
                Dim col = index Mod 2, row = index \ 2
                Dim x = left + col * (cellW + gap), y = top + row * (cellH + gap)
                AddCard(plan, $"matrix-{index + 1}", items(index), x, y, cellW, cellH, t, index)
            Next
            AddText(plan, "matrix-y", FirstNonEmpty(spec.YAxisLabel, "Y AXIS"), 20, top + 84, 88, 44, t.CaptionSize, t.TextSecondary, True, "center")
            AddText(plan, "matrix-x", FirstNonEmpty(spec.XAxisLabel, "X AXIS") & " →", w - 230, h - 42, 160, 20, t.CaptionSize, t.TextSecondary, True, "right")
        End Sub

        Private Shared Sub BuildQuote(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single, h As Single)
            AddText(plan, "quote-mark", ChrW(&H201C).ToString(), 54, 35, 120, 110, 86, t.Primary, True)
            AddText(plan, "quote", FirstNonEmpty(spec.KeyMessage, spec.Body, spec.Title), 95, 120, w - 190, 210, t.DisplaySize - 4, t.TextPrimary, True, "center")
            AddLine(plan, "quote-rule", w * 0.42F, 370, w * 0.58F, 370, t.Primary, 2)
            AddText(plan, "quote-source", FirstNonEmpty(spec.Source, spec.Subtitle, spec.Eyebrow), 150, 390, w - 300, 35, t.BodySize, t.TextSecondary, False, "center")
        End Sub

        Private Shared Sub BuildClosing(plan As SlideRenderPlan, spec As SlideDesignSpec, t As DesignTokens, w As Single, h As Single)
            AddRect(plan, "closing-panel", w * 0.74F, 0, w * 0.26F, h, t.Surface, collision:=False)
            AddRect(plan, "closing-accent", w * 0.74F, 0, 8, h, t.Primary, collision:=False)
            AddText(plan, "closing-eyebrow", spec.Eyebrow, 60, 80, 260, 24, t.CaptionSize, t.Primary, True)
            AddText(plan, "closing-title", spec.Title, 60, 145, w * 0.68F, 115, t.DisplaySize + 5, t.TextPrimary, True)
            AddText(plan, "closing-message", FirstNonEmpty(spec.KeyMessage, spec.Subtitle), 62, 285, w * 0.58F, 90, t.BodySize + 2, t.TextSecondary, False)
            If Not String.IsNullOrWhiteSpace(spec.Cta) Then
                AddRect(plan, "closing-cta", 62, 405, 210, 48, t.Primary, t.Primary)
                AddText(plan, "closing-cta-text", spec.Cta, 76, 417, 182, 24, t.CaptionSize + 1, t.Background, True, "center")
            Else
                AddRect(plan, "closing-rule", 62, 414, 110, 6, t.Primary, collision:=False)
            End If
            AddText(plan, "closing-source", spec.Source, 62, h - 52, w * 0.55F, 18, 9.5F, t.TextSecondary, False)
        End Sub
    End Class

End Namespace
