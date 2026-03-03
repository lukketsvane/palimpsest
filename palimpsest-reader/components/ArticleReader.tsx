'use client';

import { motion, useScroll, useSpring, AnimatePresence } from 'motion/react';
import { BookOpen, ChevronLeft, Share, Bookmark, List, X, FileText } from 'lucide-react';
import { useState, useEffect } from 'react';
import Markdown from 'react-markdown';
import { articleContent } from '@/lib/content';

export default function ArticleReader() {
  const { scrollYProgress } = useScroll();
  const scaleX = useSpring(scrollYProgress, {
    stiffness: 100,
    damping: 30,
    restDelta: 0.001
  });

  const [isTocOpen, setIsTocOpen] = useState(false);
  const [activeSection, setActiveSection] = useState('');

  useEffect(() => {
    const handleScroll = () => {
      const sections = articleContent.sections.map(s => document.getElementById(s.id));
      const scrollPosition = window.scrollY + 100;

      for (let i = sections.length - 1; i >= 0; i--) {
        const section = sections[i];
        if (section && section.offsetTop <= scrollPosition) {
          setActiveSection(section.id);
          break;
        }
      }
    };

    window.addEventListener('scroll', handleScroll);
    return () => window.removeEventListener('scroll', handleScroll);
  }, []);


  // Keyboard navigation
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement) return;

      const sections = articleContent.sections.map(s => s.id);
      let currentIndex = sections.findIndex(id => id === activeSection);
      
      if (currentIndex === -1) currentIndex = 0;

      if (e.key === 'ArrowRight' || e.key === 'j') {
        e.preventDefault();
        const nextId = sections[currentIndex + 1] || sections[0];
        if (nextId) scrollToSection(nextId);
      } else if (e.key === 'ArrowLeft' || e.key === 'k') {
        e.preventDefault();
        const prevId = currentIndex > 0 ? sections[currentIndex - 1] : sections[sections.length - 1];
        if (prevId) scrollToSection(prevId);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [activeSection]);

  const scrollToSection = (id: string) => {
    const element = document.getElementById(id);
    if (element) {
      window.scrollTo({
        top: element.offsetTop - 80,
        behavior: 'smooth'
      });
    }
    setIsTocOpen(false);
  };

  return (
    <div className="min-h-screen bg-white text-black font-serif relative">
      {/* Progress Bar */}
      <motion.div
        className="fixed top-0 left-0 right-0 h-1 bg-black origin-left z-50 print:hidden"
        style={{ scaleX }}
      />

      {/* Header */}
      <header className="fixed top-0 left-0 right-0 h-14 bg-white border-b-2 border-black z-40 flex items-center justify-between px-4 sm:px-6 print:hidden">
        <div className="p-2 -ml-2 text-black flex items-center gap-1">
          <span className="text-[15px] font-bold uppercase tracking-widest font-sans">palimpsest.iverfinne.no</span>
        </div>
        
                <div className="flex items-center gap-2">
          {/* Download Buttons */}
          <button 
            onClick={() => window.print()} 
            className="p-2 text-black border-2 border-transparent hover:border-black hover:bg-black hover:text-white transition-all group"
            title="Download as PDF"
          >
            <FileText className="w-4 h-4" />
          </button>
          
          <a 
            href="/palimpsest.epub" 
            download="Palimpsest - Iver Raknes Finne.epub" 
            className="p-2 text-black border-2 border-transparent hover:border-black hover:bg-black hover:text-white transition-all group"
            title="Download as EPUB"
          >
            <BookOpen className="w-4 h-4" />
          </a>

          <div className="w-px h-6 bg-gray-300 mx-1 hidden sm:block"></div>
          <button className="p-2 text-black hover:bg-black hover:text-white transition-colors">
            <Share className="w-4 h-4" />
          </button>
          <button className="p-2 text-black hover:bg-black hover:text-white transition-colors">
            <Bookmark className="w-4 h-4" />
          </button>
          <button 
            onClick={() => setIsTocOpen(true)}
            className="p-2 text-black hover:bg-black hover:text-white transition-colors hidden sm:block"
          >
            <List className="w-4 h-4" />
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="pt-24 pb-32 px-4 sm:px-6 lg:px-12 max-w-[1400px] mx-auto print:pt-0 print:px-0">
        <motion.article
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, ease: [0.16, 1, 0.3, 1] }}
        >
          {/* Title Section */}
          <div className="mb-12 border-b-4 border-black pb-8 print:border-b-2">
            <h1 className="text-4xl sm:text-6xl md:text-7xl font-black tracking-tighter mb-4 text-black uppercase leading-none">
              {articleContent.title}
            </h1>
            <p className="text-xl sm:text-2xl text-black font-bold tracking-tight mb-3">
              {articleContent.subtitle}
            </p>
            <p className="text-lg text-black italic mb-8">
              {articleContent.subSubtitle}
            </p>
            <div className="flex items-center gap-4 text-xs text-black font-bold uppercase tracking-widest font-sans border-t-2 border-black pt-4">
              <span>{articleContent.author}</span>
              <span className="w-1.5 h-1.5 bg-black" />
              <span>{articleContent.date}</span>
              <span className="w-1.5 h-1.5 bg-black" />
              <span>5 min read</span>
            </div>
          </div>

          {/* Hero Images */}
          {articleContent.hero && (
            <div className="mb-12 border-b-4 border-black pb-12 prose prose-sm sm:prose-base prose-black max-w-none prose-img:w-full prose-img:my-8 prose-p:text-sm prose-p:text-gray-600 prose-p:italic print:border-b-2">
              <Markdown>{articleContent.hero}</Markdown>
            </div>
          )}

          {/* Abstracts */}
          <div className="mb-12 border-b-4 border-black pb-12 print:border-b-2">
            {articleContent.abstracts.map((abstract, idx) => (
              <div key={idx} className="bg-white">
                <h3 className="text-xs font-black tracking-widest uppercase text-black mb-4 font-sans border-b-2 border-black pb-2">
                  {abstract.title}
                </h3>
                <p className="text-black leading-relaxed mb-6 text-sm">
                  {abstract.content}
                </p>
                <p className="text-[10px] text-black font-sans border-t border-black pt-4">
                  <span className="font-bold uppercase tracking-wider">Keywords: </span>
                  {abstract.keywords}
                </p>
              </div>
            ))}
          </div>

          {/* Sections */}
          <div className="flex flex-col gap-16 print:gap-12">
            {articleContent.sections.map((section) => (
              <motion.section 
                key={section.id} 
                id={section.id} 
                className="scroll-mt-24 break-inside-avoid mb-12"
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true, margin: "-100px" }}
                transition={{ duration: 0.5, ease: "easeOut" }}
              >
                <h2 className="text-xl sm:text-2xl font-black tracking-tight mb-4 text-black uppercase border-b-2 border-black pb-2 font-sans">
                  {section.title}
                </h2>
                <div className="prose prose-sm sm:prose-base prose-black max-w-none prose-p:leading-relaxed prose-p:text-black prose-headings:font-sans prose-headings:font-black prose-headings:uppercase prose-blockquote:border-l-4 prose-blockquote:border-black prose-blockquote:bg-gray-100 prose-blockquote:py-4 prose-blockquote:px-6 prose-blockquote:not-italic prose-blockquote:text-black prose-img:w-full prose-img:my-8 prose-a:text-black prose-a:font-bold prose-a:underline prose-a:decoration-2 prose-a:underline-offset-4 hover:prose-a:bg-black hover:prose-a:text-white columns-1 md:columns-2 gap-12 print:columns-1">
                  <Markdown>{section.content}</Markdown>
                </div>
              </motion.section>
            ))}
          </div>
        </motion.article>

        {/* Footer */}
        <footer className="mt-24 border-t-4 border-black print:hidden" />
      </main>

      {/* Floating TOC Button (Mobile) */}
      <button
        onClick={() => setIsTocOpen(true)}
        className="fixed bottom-6 right-6 p-4 bg-black text-white sm:hidden z-40 hover:bg-white hover:text-black hover:border-2 hover:border-black transition-colors print:hidden"
      >
        <List className="w-6 h-6" />
      </button>

      {/* Table of Contents Overlay */}
      <AnimatePresence>
        {isTocOpen && (
          <div className="fixed inset-0 z-50 flex justify-end print:hidden">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsTocOpen(false)}
              className="absolute inset-0 bg-white/90 backdrop-blur-sm"
            />
            <motion.div
              initial={{ x: '100%' }}
              animate={{ x: 0 }}
              exit={{ x: '100%' }}
              transition={{ duration: 0.2, ease: "easeOut" }}
              className="relative w-full max-w-sm bg-white h-full border-l-4 border-black flex flex-col"
            >
              <div className="p-4 border-b-4 border-black flex items-center justify-between bg-white">
                <h3 className="font-black text-xl tracking-widest uppercase font-sans">Contents</h3>
                <button 
                  onClick={() => setIsTocOpen(false)}
                  className="p-2 text-black hover:bg-black hover:text-white transition-colors"
                >
                  <X className="w-6 h-6" />
                </button>
              </div>
              <div className="flex-1 overflow-y-auto p-4">
                <nav className="space-y-2">
                  {articleContent.sections.map((section) => (
                    <button
                      key={section.id}
                      onClick={() => scrollToSection(section.id)}
                      className={`w-full text-left px-4 py-3 text-sm font-sans font-bold uppercase tracking-wider transition-colors border-2 ${
                        activeSection === section.id 
                          ? 'bg-black text-white border-black' 
                          : 'bg-white text-black border-transparent hover:border-black'
                      }`}
                    >
                      {section.title}
                    </button>
                  ))}
                  <div className="pt-4 mt-4 border-t-2 border-black flex flex-col gap-2">
                    <button 
                      onClick={() => { setIsTocOpen(false); setTimeout(() => window.print(), 300); }}
                      className="w-full text-left px-4 py-3 text-sm font-sans font-bold uppercase tracking-wider transition-colors border-2 bg-white text-black hover:bg-black hover:text-white flex items-center gap-2"
                    >
                      <FileText className="w-4 h-4" /> Download PDF
                    </button>
                    <a 
                      href="/palimpsest.epub" 
                      download="Palimpsest - Iver Raknes Finne.epub"
                      className="w-full text-left px-4 py-3 text-sm font-sans font-bold uppercase tracking-wider transition-colors border-2 bg-white text-black hover:bg-black hover:text-white flex items-center gap-2"
                    >
                      <BookOpen className="w-4 h-4" /> Download EPUB
                    </a>
                  </div>
                </nav>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
