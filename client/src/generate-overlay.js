import fs from 'fs';
import path from 'path';

export const svgContent = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1920 1080" width="1920" height="1080">
  <defs>
    <!-- Deep dark navy gradient for the right dashboard panel -->
    <linearGradient id="navyBg" x1="20%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#08182B" />
      <stop offset="45%" stop-color="#061527" />
      <stop offset="100%" stop-color="#040C16" />
    </linearGradient>

    <!-- Radiant golden-yellow gradient for the angled ribbon divider -->
    <linearGradient id="goldRibbon" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FFE066" />
      <stop offset="25%" stop-color="#FFC400" />
      <stop offset="75%" stop-color="#FFA800" />
      <stop offset="100%" stop-color="#FF9000" />
    </linearGradient>

    <!-- Top facet golden gradient for 3D depth -->
    <linearGradient id="facetGold" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FFD54F" />
      <stop offset="50%" stop-color="#FFB300" />
      <stop offset="100%" stop-color="#E68A00" />
    </linearGradient>

    <!-- Subtle golden glow filter -->
    <filter id="ribbonGlow" x="-30%" y="-30%" width="160%" height="160%">
      <feGaussianBlur stdDeviation="12" result="blur" />
      <feComposite in="SourceGraphic" in2="blur" operator="over" />
    </filter>

    <!-- Light streak gradients -->
    <linearGradient id="streakGrad1" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#FFE066" stop-opacity="0.95" />
      <stop offset="35%" stop-color="#FFA800" stop-opacity="0.8" />
      <stop offset="70%" stop-color="#FF7700" stop-opacity="0.3" />
      <stop offset="100%" stop-color="#FF7700" stop-opacity="0" />
    </linearGradient>

    <linearGradient id="streakGrad2" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%" stop-color="#FFC400" stop-opacity="0.7" />
      <stop offset="40%" stop-color="#FF8800" stop-opacity="0.4" />
      <stop offset="85%" stop-color="#FF5500" stop-opacity="0.1" />
      <stop offset="100%" stop-color="#FF5500" stop-opacity="0" />
    </linearGradient>
  </defs>

  <!-- Bottom-left corner golden accent triangle -->
  <polygon points="0,920 0,1080 140,1080" fill="url(#goldRibbon)" />
  <line x1="0" y1="920" x2="140" y2="1080" stroke="#FFF7D1" stroke-width="1.8" />

  <!-- Dark navy right-hand panel (left side transparent) -->
  <polygon points="1270,0 1920,0 1920,1080 970,1080" fill="url(#navyBg)" />

  <!-- Subtle translucent location pin watermark at top-right (dark slate outline matching reference photo) -->
  <g opacity="0.6" transform="translate(1385, 45)">
    <path d="M 55 0 C 24.6 0 0 24.6 0 55 C 0 95 55 145 55 145 C 55 145 110 95 110 55 C 110 24.6 85.4 0 55 0 Z"
          fill="none" stroke="#0B2238" stroke-width="7" stroke-linecap="round" stroke-linejoin="round" />
    <circle cx="55" cy="55" r="21"
            fill="none" stroke="#0B2238" stroke-width="7" />
  </g>

  <!-- Architectural City Skyline Silhouette in bottom right -->
  <g opacity="0.16" stroke="#D4A034" stroke-width="1.4" fill="none">
    <path d="M 1200 1080 L 1200 950 L 1230 950 L 1230 980 L 1250 980 L 1250 920 L 1270 890 L 1290 920 L 1290 1080" />
    <path d="M 1290 1080 L 1290 940 L 1330 940 L 1330 900 L 1350 860 L 1370 900 L 1370 1080" />
    <path d="M 1370 1080 L 1370 960 L 1410 960 L 1410 1080" />
    <path d="M 1410 1080 L 1410 880 L 1430 850 L 1450 880 L 1450 970 L 1480 970 L 1480 1080" />
    <path d="M 1480 1080 L 1480 930 L 1520 930 L 1520 1080" />
    <path d="M 1520 1080 L 1520 910 L 1545 885 L 1570 910 L 1570 1080" />
    <path d="M 1570 1080 L 1570 950 L 1610 950 L 1610 1080" />
    <path d="M 1610 1080 L 1610 890 L 1630 870 L 1650 890 L 1650 1080" />
    <path d="M 1650 1080 L 1650 940 L 1700 940 L 1700 1080" />
    <path d="M 1700 1080 L 1700 910 L 1740 910 L 1740 1080" />
    <path d="M 1740 1080 L 1740 960 L 1800 960 L 1800 1080" />
    <path d="M 1800 1080 L 1800 930 L 1850 930 L 1850 1080" />
    <path d="M 1850 1080 L 1850 950 L 1920 950 L 1920 1080" />
    <line x1="1215" y1="965" x2="1215" y2="1040" stroke-dasharray="4,6" />
    <line x1="1310" y1="955" x2="1310" y2="1040" stroke-dasharray="4,6" />
    <line x1="1430" y1="895" x2="1430" y2="1030" stroke-dasharray="4,6" />
    <line x1="1630" y1="905" x2="1630" y2="1030" stroke-dasharray="4,6" />
  </g>

  <!-- Glowing light streak sweeping from the bottom of the divider to the right -->
  <g>
    <!-- Soft wide ambient glow along bottom edge -->
    <path d="M 940 1080 Q 1120 1000 1450 1035 T 1920 1055" fill="none" stroke="url(#streakGrad2)" stroke-width="26" opacity="0.35" />
    <!-- Main intense light stream -->
    <path d="M 945 1080 Q 1110 1010 1420 1040 T 1920 1058" fill="none" stroke="url(#streakGrad1)" stroke-width="7" />
    <!-- Sharp crisp white-gold core filament -->
    <path d="M 950 1080 Q 1105 1015 1400 1042 T 1920 1060" fill="none" stroke="#FFF5CC" stroke-width="2.5" opacity="0.95" />
    <!-- Secondary lower arc -->
    <path d="M 965 1080 Q 1130 1035 1450 1055 T 1920 1070" fill="none" stroke="url(#streakGrad1)" stroke-width="3" opacity="0.7" />
  </g>

  <!-- Top angled golden facet / wing pointing down-left -->
  <polygon points="980,0 1280,0 1150,175" fill="url(#facetGold)" />
  <line x1="980" y1="0" x2="1150" y2="175" stroke="#FFF7D1" stroke-width="1.8" />

  <!-- The Angled Golden-Yellow Ribbon Divider -->
  <!-- Outer golden edge glow -->
  <polygon points="1235,-10 1275,-10 975,1090 935,1090" fill="#FFC400" opacity="0.25" filter="url(#ribbonGlow)" />
  <!-- Main sharp ribbon -->
  <polygon points="1242,0 1272,0 972,1080 942,1080" fill="url(#goldRibbon)" />
  <!-- Bright highlight pin-stripe on the left border of the ribbon -->
  <line x1="1243" y1="0" x2="943" y2="1080" stroke="#FFF7D1" stroke-width="2.2" />
</svg>
`;

const targetPath = path.resolve(process.cwd(), 'client/public/assets/ppt-slide-overlay.svg');
fs.writeFileSync(targetPath, svgContent.trim(), 'utf8');
console.log('Saved updated SVG overlay to:', targetPath);
