// Draws one slide theme's gallery picture with the bundled converter's own libraries.
//
// WHY this exists (v0.1.37): PowerPoint's Design tab lists the standard slide themes from
// sdkjs/slide/themes/theme<N>/theme.bin, themes.js (their names) and one picture strip per screen
// scale (sdkjs/common/Images/themes_thumbnail*.png). OnlyOffice makes all three with its
// `allthemesgen` tool, which euro-office-lite never runs or ships — so its editor (and ours) had an
// empty theme gallery. build/gen-themes.mjs makes theme.bin with x2t (exactly as allthemesgen
// does) and calls this program for the pictures: it repeats allthemesgen's drawing steps
// (OnlyOffice core, DesktopEditor/allthemesgen/main.cpp) against the converter's shared libraries,
// so the pictures are the editor's own drawing of each theme, not an imitation.
//
// WHY declarations instead of OnlyOffice's headers: the libraries come prebuilt from
// euro-office-lite's release; only these few exported functions are called. The declarations
// below copy the public headers' shapes (Fonts.h, doctrenderer.h, MetafileToRenderer.h).
// Classes whose size is not ours to know are built inside a generous buffer.
//
// Usage: themethumbs <converter dir> <font data dir> <theme dir holding theme.bin and media/> <w> <h> <out.png> [...]
// (font data dir: what `x2t -create-allfonts` made — AllFonts.js and font_selection.bin naming real
// font files; the bundled AllFonts.js names bare file names, and text then draws as missing glyphs.)
// Prints the theme's name on success.
#include <string>
#include <cstdio>
#include <cstdlib>
#include <new>
#include <dirent.h>

namespace NSBase {
class CBaseRefCounter {
 protected:
  volatile int m_lRef;
 public:
  CBaseRefCounter();
  virtual ~CBaseRefCounter();
  virtual int AddRef();
  virtual int Release();
};
}  // namespace NSBase
namespace NSFonts {
class IFontsCache; class IFontList; class IApplicationFontStreams;
// WHY the full virtual order up to InitializeFromFolder: the call goes through the vtable, so the
// slots must line up with the library's (Fonts.h: IApplicationFonts).
class IApplicationFonts : public NSBase::CBaseRefCounter {
 public:
  IApplicationFonts();
  virtual ~IApplicationFonts();
  virtual IFontsCache* GetCache() = 0;
  virtual IFontList* GetList() = 0;
  virtual IApplicationFontStreams* GetStreams() = 0;
  virtual void InitializeFromFolder(std::wstring strFolder, bool bIsCheckSelection = true) = 0;
};
namespace NSApplication { IApplicationFonts* Create(); }
}  // namespace NSFonts
namespace NSDoctRenderer {
class CDoctrenderer {
  void* m_pInternal;
  char spare[256];
 public:
  CDoctrenderer(const std::wstring& sAllFontsPath);
  void LoadConfig(const std::wstring& sConfigDir, const std::wstring& sAllFontsPath);
  ~CDoctrenderer();
  bool Execute(const std::wstring& strXml, std::wstring& strError);
};
class CDocBuilder {
 public:
  static void Initialize(const wchar_t* directory);
  static void Dispose();
};
}  // namespace NSDoctRenderer
class IRenderer;
class IMetafileToRenderter {
 public:
  void SetTempDirectory(const std::wstring& sDir);
  void SetMediaDirectory(const std::wstring& sDir);
  void SetThemesDirectory(const std::wstring& sDir);
};
namespace NSOnlineOfficeBinToPdf {
class CMetafileToRenderterRaster : public IMetafileToRenderter {
 public:
  CMetafileToRenderterRaster(IRenderer* pRenderer);
  ~CMetafileToRenderterRaster();
  void SetApplication(NSFonts::IApplicationFonts* pApplication);
  void SetRasterFormat(const int& value);
  void SetSaveType(const int& value);
  void SetIsOnlyFirst(const bool& value);
  void SetRasterW(const int& value);
  void SetRasterH(const int& value);
  void SetFileName(const std::wstring& value);
  bool ConvertBuffer(unsigned char* pBuffer, long lBufferLen);
};
}  // namespace NSOnlineOfficeBinToPdf

// The build's paths are ASCII (the build folder and theme<N>); a byte-wide copy is enough.
static std::wstring W(const std::string& s) { return std::wstring(s.begin(), s.end()); }
static std::string xmlEscape(const std::string& s) {
  std::string r;
  for (char c : s) r += c == '&' ? "&amp;" : c == '<' ? "&lt;" : c == '>' ? "&gt;" : std::string(1, c);
  return r;
}

int main(int argc, char** argv) {
  if (argc < 7 || (argc - 4) % 3 != 0) {
    fprintf(stderr, "usage: themethumbs <converter dir> <font data dir> <theme dir> <w> <h> <out.png> [<w> <h> <out.png>...]\n");
    return 2;
  }
  const std::string conv = argv[1], fontData = argv[2], dir = argv[3];
  const std::wstring allFonts = W(fontData + "/AllFonts.js");
  NSFonts::IApplicationFonts* fonts = NSFonts::NSApplication::Create();
  fonts->InitializeFromFolder(W(fontData));
  NSDoctRenderer::CDocBuilder::Initialize(nullptr);

  // allthemesgen's doctrenderer task: open theme.bin (PPTT) and draw its gallery picture
  // (PPTX_THEME_THUMBNAIL = 5) as drawing commands into "<theme name>.theme" beside it.
  const std::string xml = "<Settings><SrcFileType>2</SrcFileType><DstFileType>5</DstFileType><SrcFilePath>" + xmlEscape(dir) +
      "/theme.bin</SrcFilePath><DstFilePath>" + xmlEscape(dir) + "</DstFilePath><FontsDirectory>" + xmlEscape(fontData) +
      "</FontsDirectory><ImagesDirectory>" + xmlEscape(dir) + "/media</ImagesDirectory><ThemesDirectory>" + xmlEscape(dir) +
      "</ThemesDirectory></Settings>";
  std::wstring err;
  bool ok;
  {
    NSDoctRenderer::CDoctrenderer renderer(allFonts);
    renderer.LoadConfig(W(conv), allFonts);
    ok = renderer.Execute(W(xml), err);
  }
  if (!ok || !err.empty()) {
    fprintf(stderr, "themethumbs: the converter could not draw %s\n", dir.c_str());
    return 1;
  }

  std::string name, drawing;
  if (DIR* d = opendir(dir.c_str())) {
    while (dirent* e = readdir(d)) {
      std::string n = e->d_name;
      if (n.size() > 6 && n.compare(n.size() - 6, 6, ".theme") == 0) {
        name = n.substr(0, n.size() - 6);
        drawing = dir + "/" + n;
      }
    }
    closedir(d);
  }
  if (drawing.empty()) {
    fprintf(stderr, "themethumbs: no drawing for %s\n", dir.c_str());
    return 1;
  }
  FILE* f = fopen(drawing.c_str(), "rb");
  if (!f) return 1;
  fseek(f, 0, SEEK_END);
  long len = ftell(f);
  fseek(f, 0, SEEK_SET);
  unsigned char* data = static_cast<unsigned char*>(malloc(len));
  if (!data || fread(data, 1, len, f) != static_cast<size_t>(len)) { fclose(f); return 1; }
  fclose(f);
  remove(drawing.c_str());

  // One picture per requested size, as allthemesgen's raster loop (PNG = format 4).
  for (int i = 4; i + 2 < argc; i += 3) {
    alignas(64) static char space[16384];
    auto* raster = new (space) NSOnlineOfficeBinToPdf::CMetafileToRenderterRaster(nullptr);
    raster->SetMediaDirectory(W(dir));
    raster->SetThemesDirectory(L"");
    raster->SetTempDirectory(W(dir));
    raster->SetApplication(fonts);
    raster->SetRasterFormat(4);
    raster->SetSaveType(0);
    raster->SetIsOnlyFirst(true);
    raster->SetRasterW(atoi(argv[i]));
    raster->SetRasterH(atoi(argv[i + 1]));
    raster->SetFileName(W(argv[i + 2]));
    raster->ConvertBuffer(data, len);
    raster->~CMetafileToRenderterRaster();
  }
  free(data);
  NSDoctRenderer::CDocBuilder::Dispose();
  printf("%s\n", name.c_str());
  return 0;
}
