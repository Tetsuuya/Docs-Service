/**
 * 10 DISTINCT VISUAL TEMPLATES FOR PRESENTATIONS
 * Each template defines unique positioning, colors, and styling
 */

export const VISUAL_TEMPLATES = {
  // ═══════════════════════════════════════════════════════════
  // TEMPLATE 1: MODERN DARK (Current default)
  // ═══════════════════════════════════════════════════════════
  modern_dark: {
    name: 'Modern Dark',
    hero: {
      showTopBar: true,
      topBarHeight: 0.08,
      showBadge: true,
      badgePosition: { x: 0.5, y: 0.4 },
      titlePosition: { x: 0.5, y: 0.9 },
      subtitlePosition: { x: 0.5, y: 3.3 },
      imagePosition: { x: 5.9, y: 0.9 },
      imageSize: { w: 3.7, h: 3.8 },
      showFooter: true,
      footerY: 5.1
    },
    content: {
      showTopBar: true,
      titlePosition: { x: 0.5, y: 0.3 },
      subtitlePosition: { x: 0.5, y: 0.8 },
      contentTop: 1.25,
      showFooter: true,
      footerY: 5.15,
      cardBorderWidth: 1,
      cardRadius: 0.15,
      useAccentHeaders: true
    },
    colors: {
      cardOpacity: 1,
      useGradients: false
    }
  },

  // ═══════════════════════════════════════════════════════════
  // TEMPLATE 2: MINIMAL LIGHT
  // ═══════════════════════════════════════════════════════════
  minimal_light: {
    name: 'Minimal Light',
    hero: {
      showTopBar: false,
      showBadge: false,
      titlePosition: { x: 1.0, y: 1.5 },
      subtitlePosition: { x: 1.0, y: 3.5 },
      imagePosition: { x: 6.0, y: 1.0 },
      imageSize: { w: 3.5, h: 4.0 },
      showFooter: false
    },
    content: {
      showTopBar: false,
      titlePosition: { x: 0.5, y: 0.5 },
      subtitlePosition: { x: 0.5, y: 1.0 },
      contentTop: 1.5,
      showFooter: false,
      cardBorderWidth: 0,
      cardRadius: 0.3,
      useAccentHeaders: false
    },
    colors: {
      overrideBackground: 'FAFAFA',
      overrideCardBg: 'FFFFFF',
      overridePrimary: '1A1A1A',
      overrideSecondary: '666666',
      cardOpacity: 0.95,
      useGradients: false
    }
  },

  // ═══════════════════════════════════════════════════════════
  // TEMPLATE 3: BOLD CORPORATE
  // ═══════════════════════════════════════════════════════════
  bold_corporate: {
    name: 'Bold Corporate',
    hero: {
      showTopBar: false,
      showBadge: true,
      badgePosition: { x: 0.3, y: 0.3 },
      titlePosition: { x: 0.3, y: 1.0 },
      subtitlePosition: { x: 0.3, y: 2.8 },
      imagePosition: { x: 5.5, y: 0.5 },
      imageSize: { w: 4.0, h: 4.5 },
      showFooter: true,
      footerY: 5.2
    },
    content: {
      showTopBar: false,
      titlePosition: { x: 0.3, y: 0.3 },
      subtitlePosition: { x: 0.3, y: 0.8 },
      contentTop: 1.3,
      showFooter: true,
      footerY: 5.2,
      cardBorderWidth: 3,
      cardRadius: 0.05,
      useAccentHeaders: true,
      leftAligned: true
    },
    colors: {
      cardOpacity: 1,
      useGradients: false,
      heavyBorders: true
    }
  },

  // ═══════════════════════════════════════════════════════════
  // TEMPLATE 4: CREATIVE ASYMMETRIC
  // ═══════════════════════════════════════════════════════════
  creative_asymmetric: {
    name: 'Creative Asymmetric',
    hero: {
      showTopBar: false,
      showBadge: true,
      badgePosition: { x: 6.5, y: 0.3 },
      titlePosition: { x: 1.5, y: 1.2 },
      subtitlePosition: { x: 2.0, y: 3.0 },
      imagePosition: { x: 0.3, y: 0.8 },
      imageSize: { w: 3.2, h: 3.5 },
      showFooter: false
    },
    content: {
      showTopBar: false,
      titlePosition: { x: 1.0, y: 0.4 },
      subtitlePosition: { x: 1.5, y: 0.9 },
      contentTop: 1.4,
      showFooter: false,
      cardBorderWidth: 2,
      cardRadius: 0.25,
      useAccentHeaders: false,
      asymmetricLayout: true
    },
    colors: {
      cardOpacity: 0.9,
      useGradients: true
    }
  },

  // ═══════════════════════════════════════════════════════════
  // TEMPLATE 5: CLASSIC CENTERED
  // ═══════════════════════════════════════════════════════════
  classic_centered: {
    name: 'Classic Centered',
    hero: {
      showTopBar: false,
      showBadge: false,
      titlePosition: { x: 2.0, y: 1.8 },
      subtitlePosition: { x: 2.0, y: 3.2 },
      imagePosition: { x: 3.0, y: 0.4 },
      imageSize: { w: 4.0, h: 1.2 },
      showFooter: true,
      footerY: 5.3,
      centered: true
    },
    content: {
      showTopBar: false,
      titlePosition: { x: 2.0, y: 0.4 },
      subtitlePosition: { x: 2.5, y: 0.9 },
      contentTop: 1.4,
      showFooter: true,
      footerY: 5.3,
      cardBorderWidth: 1,
      cardRadius: 0.1,
      useAccentHeaders: false,
      centered: true
    },
    colors: {
      cardOpacity: 0.95,
      useGradients: false
    }
  },

  // ═══════════════════════════════════════════════════════════
  // TEMPLATE 6: TECH GRADIENT
  // ═══════════════════════════════════════════════════════════
  tech_gradient: {
    name: 'Tech Gradient',
    hero: {
      showTopBar: true,
      topBarHeight: 0.15,
      showBadge: false,
      titlePosition: { x: 0.5, y: 1.2 },
      subtitlePosition: { x: 0.5, y: 3.5 },
      imagePosition: { x: 6.2, y: 1.0 },
      imageSize: { w: 3.3, h: 3.5 },
      showFooter: false
    },
    content: {
      showTopBar: true,
      topBarHeight: 0.12,
      titlePosition: { x: 0.5, y: 0.4 },
      subtitlePosition: { x: 0.5, y: 0.9 },
      contentTop: 1.3,
      showFooter: false,
      cardBorderWidth: 0,
      cardRadius: 0.3,
      useAccentHeaders: true
    },
    colors: {
      cardOpacity: 0.85,
      useGradients: true,
      gradientDirection: 'horizontal'
    }
  },

  // ═══════════════════════════════════════════════════════════
  // TEMPLATE 7: MAGAZINE EDITORIAL
  // ═══════════════════════════════════════════════════════════
  magazine_editorial: {
    name: 'Magazine Editorial',
    hero: {
      showTopBar: false,
      showBadge: false,
      titlePosition: { x: 0.4, y: 3.5 },
      subtitlePosition: { x: 0.4, y: 4.5 },
      imagePosition: { x: 0.3, y: 0.3 },
      imageSize: { w: 9.4, h: 3.0 },
      showFooter: false,
      imageFirst: true
    },
    content: {
      showTopBar: false,
      titlePosition: { x: 0.4, y: 0.2 },
      subtitlePosition: { x: 0.4, y: 0.7 },
      contentTop: 1.1,
      showFooter: false,
      cardBorderWidth: 0,
      cardRadius: 0,
      useAccentHeaders: false,
      magazineStyle: true
    },
    colors: {
      cardOpacity: 0.9,
      useGradients: false
    }
  },

  // ═══════════════════════════════════════════════════════════
  // TEMPLATE 8: ACADEMIC GRID
  // ═══════════════════════════════════════════════════════════
  academic_grid: {
    name: 'Academic Grid',
    hero: {
      showTopBar: false,
      showBadge: true,
      badgePosition: { x: 0.5, y: 0.5 },
      titlePosition: { x: 0.5, y: 1.5 },
      subtitlePosition: { x: 0.5, y: 3.0 },
      imagePosition: { x: 5.8, y: 1.5 },
      imageSize: { w: 3.7, h: 3.0 },
      showFooter: true,
      footerY: 5.0
    },
    content: {
      showTopBar: false,
      titlePosition: { x: 0.5, y: 0.4 },
      subtitlePosition: { x: 0.5, y: 0.85 },
      contentTop: 1.2,
      showFooter: true,
      footerY: 5.0,
      cardBorderWidth: 1.5,
      cardRadius: 0,
      useAccentHeaders: false,
      gridBased: true
    },
    colors: {
      cardOpacity: 1,
      useGradients: false,
      monochromeMode: true
    }
  },

  // ═══════════════════════════════════════════════════════════
  // TEMPLATE 9: LUXURY DARK
  // ═══════════════════════════════════════════════════════════
  luxury_dark: {
    name: 'Luxury Dark',
    hero: {
      showTopBar: false,
      showBadge: true,
      badgePosition: { x: 3.5, y: 0.3 },
      titlePosition: { x: 1.0, y: 1.5 },
      subtitlePosition: { x: 1.0, y: 3.8 },
      imagePosition: { x: 5.5, y: 1.0 },
      imageSize: { w: 4.0, h: 4.0 },
      showFooter: true,
      footerY: 5.2
    },
    content: {
      showTopBar: false,
      titlePosition: { x: 1.0, y: 0.5 },
      subtitlePosition: { x: 1.0, y: 1.0 },
      contentTop: 1.6,
      showFooter: true,
      footerY: 5.2,
      cardBorderWidth: 2,
      cardRadius: 0.2,
      useAccentHeaders: true
    },
    colors: {
      overrideBackground: '0A0A0A',
      overrideCardBg: '1A1A1A',
      overrideAccent: 'D4AF37',
      cardOpacity: 1,
      useGradients: false,
      luxuryMode: true
    }
  },

  // ═══════════════════════════════════════════════════════════
  // TEMPLATE 10: BRUTALIST
  // ═══════════════════════════════════════════════════════════
  brutalist: {
    name: 'Brutalist',
    hero: {
      showTopBar: false,
      showBadge: false,
      titlePosition: { x: 0.2, y: 0.2 },
      subtitlePosition: { x: 0.2, y: 2.5 },
      imagePosition: { x: 5.0, y: 0.2 },
      imageSize: { w: 4.8, h: 5.4 },
      showFooter: false
    },
    content: {
      showTopBar: false,
      titlePosition: { x: 0.2, y: 0.2 },
      subtitlePosition: { x: 0.2, y: 0.7 },
      contentTop: 1.1,
      showFooter: false,
      cardBorderWidth: 4,
      cardRadius: 0,
      useAccentHeaders: false,
      brutalist: true
    },
    colors: {
      overrideBackground: 'FFFFFF',
      overrideCardBg: 'FFFFFF',
      overridePrimary: '000000',
      overrideSecondary: '000000',
      overrideAccent: '000000',
      cardOpacity: 1,
      useGradients: false,
      highContrast: true
    }
  }
};

/**
 * Get a random template or specific template by name
 */
export function getTemplate(templateName = null) {
  if (templateName && VISUAL_TEMPLATES[templateName]) {
    return { name: templateName, ...VISUAL_TEMPLATES[templateName] };
  }
  
  // Random selection
  const templateNames = Object.keys(VISUAL_TEMPLATES);
  const randomName = templateNames[Math.floor(Math.random() * templateNames.length)];
  return { name: randomName, ...VISUAL_TEMPLATES[randomName] };
}
