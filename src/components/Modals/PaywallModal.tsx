import React, { useState } from 'react';
import {
  Crown,
  Check,
  X,
  Clock,
  Table,
  FileDown,
  Type,
  Maximize2,
  Link,
  ShieldCheck,
  Sparkles,
  RefreshCw,
  Sliders,
} from 'lucide-react';
import { GatedFeature, Language, PlanTier } from '../../types/office';
import { getTranslation } from '../../i18n/translations';

interface PaywallModalProps {
  isOpen: boolean;
  onClose: () => void;
  isTrialExpired: boolean;
  lang: Language;
  gatedFeature?: GatedFeature | null;
  onUpgrade: () => void;
  currentPlan: PlanTier;
  onTogglePlan: (plan: PlanTier) => void;
  onResetTimer?: () => void;
}

export const PaywallModal: React.FC<PaywallModalProps> = ({
  isOpen,
  onClose,
  isTrialExpired,
  lang,
  gatedFeature,
  onUpgrade,
  currentPlan,
  onTogglePlan,
  onResetTimer,
}) => {
  const [isProcessing, setIsProcessing] = useState(false);
  const [isSuccess, setIsSuccess] = useState(false);
  const t = getTranslation(lang);

  if (!isOpen) return null;

  const handleUpgradeClick = () => {
    setIsProcessing(true);
    setTimeout(() => {
      setIsProcessing(false);
      setIsSuccess(true);
      setTimeout(() => {
        setIsSuccess(false);
        onUpgrade();
        onClose();
      }, 1000);
    }, 800);
  };

  const featureList: { key: GatedFeature; icon: React.ReactNode; title: string }[] = [
    { key: 'unlimited_time', icon: <Clock className="w-4 h-4 text-amber-500" />, title: t.featUnlimitedTime },
    { key: 'tables', icon: <Table className="w-4 h-4 text-emerald-500" />, title: t.featTables },
    { key: 'save_export', icon: <FileDown className="w-4 h-4 text-blue-500" />, title: t.featSaveExport },
    { key: 'font_family', icon: <Type className="w-4 h-4 text-purple-500" />, title: t.featFontFamily },
    { key: 'text_size', icon: <Sliders className="w-4 h-4 text-indigo-500" />, title: t.featTextSize },
    { key: 'page_size', icon: <Maximize2 className="w-4 h-4 text-rose-500" />, title: t.featPageSize },
    { key: 'hyperlinks', icon: <Link className="w-4 h-4 text-sky-500" />, title: t.featHyperlinks },
  ];

  return (
    <div
      className={`fixed inset-0 z-50 flex items-center justify-center p-4 transition-all duration-200 ${
        isTrialExpired ? 'bg-black/80 backdrop-blur-md' : 'bg-black/50 backdrop-blur-xs'
      }`}
    >
      <div
        className="w-full max-w-xl bg-white dark:bg-[#201f1e] text-gray-900 dark:text-gray-100 rounded-2xl shadow-2xl border border-amber-300 dark:border-amber-600/40 overflow-hidden relative animate-in zoom-in-95 duration-200"
      >
        {/* Top Gold Accent Bar */}
        <div className="h-2 bg-gradient-to-r from-amber-400 via-yellow-400 to-amber-500" />

        {/* Close Button (disabled/hidden if trial expired to freeze canvas) */}
        {!isTrialExpired && (
          <button
            onClick={onClose}
            className="absolute top-4 right-4 rtl:right-auto rtl:left-4 p-1.5 rounded-full hover:bg-gray-100 dark:hover:bg-[#323130] text-gray-400 hover:text-gray-600 dark:hover:text-gray-200 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        )}

        {/* Modal Header */}
        <div className="p-6 md:p-8 text-center space-y-3">
          <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-gradient-to-tr from-amber-500 to-yellow-300 text-white shadow-lg shadow-amber-500/20 mb-1">
            <Crown className="w-8 h-8" />
          </div>

          <div>
            <div className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-amber-100 dark:bg-amber-950/60 text-amber-800 dark:text-amber-300 font-bold text-xs uppercase tracking-wider mb-2">
              <Sparkles className="w-3.5 h-3.5" />
              <span>Office 365+ Plan</span>
            </div>

            <h2 className="text-2xl md:text-3xl font-extrabold text-gray-900 dark:text-white tracking-tight">
              {isTrialExpired ? t.trialEnded : t.upgradeModalTitle}
            </h2>

            <p className="text-sm text-gray-600 dark:text-gray-300 mt-1 max-w-md mx-auto">
              {isTrialExpired ? t.trialEndedMsg : t.upgradeModalSubtitle}
            </p>
          </div>

          {/* Pricing Highlight Pill */}
          <div className="p-3 rounded-xl bg-gradient-to-r from-amber-500/10 via-yellow-500/10 to-amber-500/10 border border-amber-300/60 dark:border-amber-500/30 flex items-center justify-between max-w-sm mx-auto">
            <div className="text-left rtl:text-right">
              <div className="text-xs font-semibold text-gray-500 dark:text-gray-400 uppercase">
                {t.planPro}
              </div>
              <div className="text-xs text-amber-700 dark:text-amber-400 font-medium">
                {t.unlimitedSession}
              </div>
            </div>
            <div className="text-right rtl:text-left">
              <div className="text-2xl font-black text-gray-900 dark:text-white">
                {lang === 'ar' ? '89.99$ ' : '$89.99 '}
                <span className="text-xs font-normal text-gray-500">/ mo</span>
              </div>
            </div>
          </div>
        </div>

        {/* Features Checklist */}
        <div className="px-6 md:px-8 pb-4">
          <div className="text-xs font-semibold uppercase tracking-wider text-gray-500 dark:text-gray-400 mb-2">
            What&apos;s included in Office 365+:
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-2 text-xs">
            {featureList.map((f) => {
              const isTargeted = gatedFeature === f.key;
              return (
                <div
                  key={f.key}
                  className={`flex items-start gap-2.5 p-2 rounded-lg transition-colors ${
                    isTargeted
                      ? 'bg-amber-100 dark:bg-amber-950/40 border border-amber-400 font-semibold'
                      : 'bg-gray-50 dark:bg-[#1b1a19] border border-gray-100 dark:border-[#323130]'
                  }`}
                >
                  <div className="mt-0.5 shrink-0">{f.icon}</div>
                  <span className="leading-snug text-gray-800 dark:text-gray-200">
                    {f.title}
                  </span>
                </div>
              );
            })}
          </div>
        </div>

        {/* Upgrade Call-to-action */}
        <div className="p-6 md:px-8 bg-gray-50 dark:bg-[#181716] border-t border-gray-200 dark:border-[#323130] space-y-3">
          <button
            onClick={handleUpgradeClick}
            disabled={isProcessing || isSuccess}
            className="w-full py-3.5 px-6 rounded-xl bg-gradient-to-r from-amber-500 to-amber-600 hover:from-amber-600 hover:to-amber-700 text-white font-bold text-sm md:text-base shadow-lg shadow-amber-500/25 transition-all flex items-center justify-center gap-2 cursor-pointer disabled:opacity-75"
          >
            {isProcessing ? (
              <>
                <RefreshCw className="w-5 h-5 animate-spin" />
                <span>Activating Office 365+...</span>
              </>
            ) : isSuccess ? (
              <>
                <Check className="w-5 h-5" />
                <span>{t.subscriptionSuccess}</span>
              </>
            ) : (
              <>
                <Crown className="w-5 h-5" />
                <span>{t.upgradeNow}</span>
              </>
            )}
          </button>

          <div className="flex items-center justify-center gap-2 text-[11px] text-gray-500 dark:text-gray-400">
            <ShieldCheck className="w-3.5 h-3.5 text-emerald-500" />
            <span>Instant activation • Cancel anytime • 100% Mock demo persistence</span>
          </div>

          {/* Evaluation / Test Mode Switcher */}
          <div className="pt-3 border-t border-gray-200 dark:border-[#323130] flex items-center justify-between text-xs">
            <span className="font-semibold text-gray-500 dark:text-gray-400">
              🛠️ {t.demoToggle}:
            </span>
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => {
                  onTogglePlan('free');
                  if (onResetTimer) onResetTimer();
                  onClose();
                }}
                className={`px-2.5 py-1 rounded text-xs transition-colors cursor-pointer ${
                  currentPlan === 'free'
                    ? 'bg-gray-300 dark:bg-gray-700 font-bold text-gray-900 dark:text-white'
                    : 'hover:bg-gray-200 dark:hover:bg-[#323130] text-gray-600 dark:text-gray-300'
                }`}
              >
                Free Plan
              </button>
              <button
                type="button"
                onClick={() => {
                  onTogglePlan('pro');
                  onClose();
                }}
                className={`px-2.5 py-1 rounded text-xs transition-colors cursor-pointer ${
                  currentPlan === 'pro'
                    ? 'bg-amber-500 text-white font-bold'
                    : 'hover:bg-amber-100 dark:hover:bg-amber-950/40 text-amber-700 dark:text-amber-400 font-semibold'
                }`}
              >
                Office 365+ (PRO)
              </button>
              {onResetTimer && (
                <button
                  type="button"
                  onClick={() => {
                    onResetTimer();
                    onClose();
                  }}
                  className="px-2 py-1 rounded hover:bg-gray-200 dark:hover:bg-[#323130] text-gray-500 cursor-pointer text-[11px]"
                  title="Reset 10-min countdown timer"
                >
                  ↻ Reset 10m
                </button>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
