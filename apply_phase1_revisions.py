#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
阶段1修订脚本：理论框架与叙事强化
"""

import re

def apply_phase1_revisions(text):
    """应用第1阶段修订"""

    # 修订1：Abstract重写（新版本替换）
    abstract_old_pattern = (
        r"Cities are increasingly governed through digital networks.*?"
        r"efficiency-oriented operation\.Keywords:"
    )

    abstract_new = (
        "Cities are increasingly governed through digital networks as well as physical infrastructure, "
        "but the environmental implications of this coordination are not automatically positive. This study "
        "incorporates digital space as a functional dimension alongside production, living, and ecological functions, "
        "and examines how multidimensional spatial coordination affects carbon emissions in 284 Chinese cities from 2006 to 2024. "
        "The results reveal an inverted U-shaped relationship between multidimensional coordination and emissions—a coordination paradox. "
        "While coordination can reduce information frictions and improve allocation efficiency, early-stage coordination may simultaneously "
        "intensify infrastructure expansion, industrial reorganization, and energy demand. Most sample cities remain below the turning point (0.542), "
        "indicating that the efficiency-oriented stage is not yet dominant. Restructuring analysis suggests that industrial upgrading and green innovation "
        "function as stage-dependent pathways rather than immediate emission-reduction mechanisms. Digital economy development attenuates the marginal "
        "emission cost of coordination and flattens the nonlinear pattern. The core finding is a governance tradeoff: multidimensional coordination can "
        "only support sustainable urban transition when data-enabled governance mechanisms shift cities from expansion-oriented integration toward "
        "efficiency-oriented operation. Sequencing matters.Keywords:"
    )

    text = re.sub(abstract_old_pattern, abstract_new, text, flags=re.DOTALL)

    # 修订2：强化协调悖论核心段落
    paradox_old = (
        r"This creates a central tension for sustainable-city research\. "
        r"Spatial coordination can reduce.*?"
        r"efficiency-oriented governance\."
    )

    paradox_new = (
        "This creates a central tension for sustainable-city research. Spatial coordination can "
        "reduce coordination frictions and improve resource allocation. Yet in contexts with expanding infrastructure demand, "
        "weak energy substitution, and limited governance capacity, the same coordination mechanisms can intensify infrastructure expansion, "
        "functional reorganization, industrial agglomeration, and energy intensity. This coordination paradox is not incidental. It reflects "
        "a fundamental governance sequencing problem: whether cities coordinate through expansion-oriented integration (building new infrastructure, "
        "reorganizing existing sectors, and scaling up activity) or through efficiency-oriented governance (optimizing existing systems, improving matching, "
        "and reducing redundancy). The environmental outcome depends not on coordination itself but on which governance pathway coordination serves."
    )

    text = re.sub(paradox_old, paradox_new, text, flags=re.DOTALL)

    # 修订3：Digital space强化段落
    digital_insert_after = "internet-related energy demand, and rebound effects"
    digital_new_para = (
        " Critically, digital infrastructure and platforms are not external tools applied to physical cities. "
        "Digital space functions as a constitutive urban spatial dimension that reorganizes production networks, living services, and "
        "ecological monitoring. Treating digital space as a fourth functional dimension—alongside production, living, and ecological spaces—"
        "clarifies how digital governance shapes the fundamental organization and efficiency of urban systems. This conceptual reframing is "
        "essential for understanding why coordination through digital platforms can either amplify or attenuate emission trajectories."
    )

    if "internet-related energy demand, and rebound effects" in text:
        text = text.replace(
            "internet-related energy demand, and rebound effects,",
            "internet-related energy demand, and rebound effects." + digital_new_para + " Yet digitalization"
        )

    # 修订4：改进H1假设叙述
    h1_old = "H1. Urban multidimensional spatial coupling coordination exerts an inverted U-shaped effect on carbon emissions."

    h1_new = (
        "H1. Urban multidimensional spatial coupling coordination exhibits an inverted U-shaped association with "
        "carbon emissions because governance tradeoffs shape which mechanisms dominate. At low coordination levels, integration intensifies "
        "infrastructure expansion and energy demand (expansion effect). As coordination deepens, efficiency improvements and improved allocation "
        "become more visible (efficiency effect). The environmental outcome reflects which governance mode dominates: expansion-oriented coordination "
        "(dominates below the turning point) or efficiency-oriented coordination (emerges above the turning point)."
    )

    text = text.replace(h1_old, h1_new)

    # 修订5：强化turning point解释
    turning_point_old = (
        r"The estimated turning point is 0\.542,.*?"
        r"and the nlcom 95% confidence interval for the turning point is \[0\.450, 0\.634\]\."
    )

    turning_point_new = (
        "The estimated turning point is 0.542. Critically, most sample cities remain substantially below this point, "
        "indicating that the transition from expansion-oriented to efficiency-oriented governance has not yet occurred for most cities. "
        "This is not a deficiency in coordination itself but a reflection of development stage. Cities below the turning point remain in a phase "
        "where infrastructure expansion, sectoral reorganization, and agglomeration dominate, generating construction and energy costs that outweigh "
        "coordination benefits. The formal nonlinear test confirms the inverted-U pattern: left-endpoint marginal effect is significantly positive "
        "(6.502, p < 0.001), right-endpoint is significantly negative (-4.325, p = 0.005), and the 95% confidence interval is [0.450, 0.634]. "
        "The implication is governance-sequencing: early coordination should focus on preparing efficiency-oriented mechanisms (data infrastructure, "
        "digital public goods, performance monitoring) rather than accelerating expansion-oriented integration."
    )

    text = re.sub(turning_point_old, turning_point_new, text, flags=re.DOTALL)

    # 修订6：Conclusion强化
    conclusion_old = (
        r"This study investigates how urban multidimensional spatial coupling coordination.*?"
        r"governance capacity\."
    )

    conclusion_new = (
        "This study demonstrates a coordination paradox in China's rapid urban transformation. Using data from 284 cities over 2006-2024 and "
        "treating digital space as a constitutive fourth functional dimension, the paper shows that multidimensional coordination exhibits an "
        "inverted U-shaped relationship with emissions. The paradox is not incidental: coordination can simultaneously reduce information frictions "
        "and intensify infrastructure expansion, industrial reorganization, and energy demand. The environmental outcome depends entirely on governance "
        "sequencing—whether coordination mechanisms serve expansion-oriented integration or efficiency-oriented operation. Empirically, most cities "
        "remain on the expansion-oriented side of the turning point (0.542), indicating that the efficiency transition has not yet become dominant."
    )

    text = re.sub(conclusion_old, conclusion_new, text, flags=re.DOTALL)

    # 修订7：政策建议第一条
    policy1_old = (
        r"First, urban low-carbon policy should prioritize governance quality.*?"
        r"ecological, and digital spaces\."
    )

    policy1_new = (
        "First, urban low-carbon policy must prioritize governance sequencing over coordination quantity. Most sample cities remain below the turning "
        "point, meaning that expanding coordination without foundational governance reform will likely amplify emissions. Policy should focus on three "
        "interlocking elements: (1) data interoperability and digital public goods that enable city-wide matching and reduce waste; (2) green energy "
        "substitution that decouples coordination from carbon intensity; (3) performance-based evaluation systems that measure emission efficiency rather "
        "than coordination breadth. These elements must advance together, not sequentially."
    )

    text = re.sub(policy1_old, policy1_new, text, flags=re.DOTALL)

    return text

# 读取原始文本
with open("d:/Workplace/SCS/manuscript_text_extracted.txt", "r", encoding="utf-8") as f:
    original_text = f.read()

# 应用修订
revised_text = apply_phase1_revisions(original_text)

# 保存修订后的文本
with open("d:/Workplace/SCS/manuscript_phase1_revised.txt", "w", encoding="utf-8") as f:
    f.write(revised_text)

print("✓ Phase 1 revisions applied")
print("✓ Saved to: manuscript_phase1_revised.txt")

# 显示修订统计
original_len = len(original_text)
revised_len = len(revised_text)
print(f"\nOriginal: {original_len} chars")
print(f"Revised: {revised_len} chars")
print(f"Difference: {revised_len - original_len:+d} chars")
