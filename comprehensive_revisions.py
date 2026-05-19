#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SCS深度修订 - 全6阶段综合修订脚本
采用分段替换策略，避免复杂正则表达式
"""

def apply_comprehensive_revisions(text):
    """应用全部6阶段修订"""

    # ========== Phase 1: 理论框架与叙事 ==========

    # 1.1 修改关键短语 - "coordination paradox"强化
    replacements = [
        # Phase 1: 理论强化
        ("The pattern reflects a coordination paradox:",
         "The pattern reflects a coordination paradox—a fundamental governance tradeoff:"),

        ("whether coordination remains expansion-oriented or becomes efficiency-oriented.",
         "whether coordination serves expansion-oriented integration (building new infrastructure and scaling activity) "
         "or efficiency-oriented operation (optimizing existing systems and reducing redundancy)."),

        ("This coordination paradox is especially relevant",
         "This coordination paradox—where integration simultaneously reduces frictions and intensifies demand—is especially relevant"),

        ("Spatial coordination can reduce coordination frictions, improve resource allocation, and support cleaner urban operation.",
         "Spatial coordination can reduce information frictions and improve resource allocation."),

        ("At the same time, it can intensify infrastructure expansion",
         "Yet in contexts with expanding infrastructure demand and limited governance capacity, it can intensify infrastructure expansion"),

        # Phase 2: Digital space强化
        ("Digital space is therefore not merely an external technological condition.",
         "Digital infrastructure and platforms are not merely external technological tools."),

        ("It functions as an urban spatial dimension because it structures flows",
         "Digital space functions as a constitutive urban spatial dimension because it structures flows"),

        # Phase 3: 机制叙述降级 - 避免因果中介语言
        ("The mechanism section must avoid causal mediation language",
         "Mechanism pathways are interpreted as restructuring associations, not causal decompositions."),

        ("transition-related pathway",
         "stage-dependent pathway"),

        ("reduction pathways",
         "transition-related pathways"),

        # Phase 4: 转折点解释强化
        ("This value lies within the SCCD range of 0.121 to 0.822, but it should not be read as a universal policy threshold.",
         "Critically, most sample cities remain substantially below this point, indicating that the transition from expansion-oriented "
         "to efficiency-oriented governance has not yet become dominant."),

        ("Most cities remain below the turning point, so the efficiency-oriented stage is not yet dominant for most of the sample.",
         "This is not a coordination deficiency but a reflection of development stage: cities below the turning point remain in a phase "
         "where infrastructure expansion, sectoral reorganization, and agglomeration generate energy costs exceeding coordination benefits."),

        # Phase 5: 语言压缩 - 缩短冗长句子
        ("This study uses panel data for 284 prefecture-level and above cities in China over the period 2006-2024",
         "This study examines 284 Chinese cities from 2006-2024"),

        ("the paper identifies a coordination paradox. Multidimensional coordination can reduce frictions",
         "the paper identifies a coordination paradox. Coordination can reduce frictions"),

        ("Its environmental consequences therefore depend on transition stage and governance capacity.",
         "Environmental outcomes depend on governance sequencing and capacity."),

        # Phase 6: Reviewer风险审计 - 保守化表述
        ("Digital economy development mainly reduces",
         "Digital economy development can attenuate"),

        ("is strongly associated with",
         "is associated with"),

        ("demonstrates that",
         "provides evidence that"),

        ("proves that coordination",
         "suggests that coordination"),

        # 其他SCS风格改进
        ("carbon effect",
         "emission implications"),

        ("carbon-related pathways",
         "restructuring pathways"),

        ("policy shock",
         "governance intervention"),

        ("directly improves low-carbon transition",
         "can support efficiency-oriented transition"),

        ("promotes sustainability directly",
         "attenuates emission costs"),

        ("low-carbon mechanisms",
         "restructuring dynamics"),

        ("indirectly reduces emissions",
         "is associated with emission-relevant restructuring"),
    ]

    # 应用所有替换
    for old, new in replacements:
        if old in text:
            text = text.replace(old, new)
            print(f"✓ Replaced: {old[:50]}...")

    return text


# 读取原始文本
print("Reading manuscript...")
with open("d:/Workplace/SCS/manuscript_text_extracted.txt", "r", encoding="utf-8") as f:
    text = f.read()

print(f"Original length: {len(text)} characters\n")

# 应用修订
print("Applying revisions...\n")
revised_text = apply_comprehensive_revisions(text)

# 保存修订版本
output_file = "d:/Workplace/SCS/manuscript_comprehensive_revised.txt"
with open(output_file, "w", encoding="utf-8") as f:
    f.write(revised_text)

print(f"\n✓ Revisions complete")
print(f"✓ New length: {len(revised_text)} characters")
print(f"✓ Saved to: {output_file}")
print(f"\nRevision summary:")
print(f"- Phase 1: Theory & narrative strengthening")
print(f"- Phase 2: Digital space conceptualization")
print(f"- Phase 3: Mechanism narrative downgrading")
print(f"- Phase 4: Turning point reinterpretation")
print(f"- Phase 5: Language compression")
print(f"- Phase 6: Reviewer risk audit")
