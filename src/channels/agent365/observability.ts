/**
 * Agent 365 channel — observability wiring.
 *
 * Configures the Microsoft.Agents A365 observability middleware on a
 * CloudAdapter (baggage propagation + per-turn output logging exported to
 * the M365 admin telemetry plane), and provides a per-turn token preload
 * helper used inside `app.onTurn('beforeTurn', ...)`.
 */

import {
  CloudAdapter,
  TurnContext,
  type Authorization,
} from '@microsoft/agents-hosting';
import {
  ObservabilityHostingManager,
  AgenticTokenCacheInstance,
} from '@microsoft/agents-a365-observability-hosting';
import { logger } from '../../logger.js';

const SUBSTRATE_SCOPE = 'https://substrate.office.com/.default';

/**
 * Attach the observability middleware to the CloudAdapter. Failures are
 * logged but never fatal — the channel still routes messages without
 * telemetry export.
 */
export function configureObservability(adapter: CloudAdapter): void {
  try {
    new ObservabilityHostingManager().configure(adapter, {
      enableBaggage: true,
      enableOutputLogging: true,
    });
  } catch (err) {
    logger.warn({ err }, 'Agent 365: observability middleware not configured');
  }
}

/**
 * Best-effort observability access-token preload. Should be called from
 * the `beforeTurn` handler so subsequent observability writes have a
 * fresh token cached.
 */
export async function preloadObservabilityToken(
  context: TurnContext,
  auth: Authorization | undefined,
  clientId: string,
  tenantId: string,
): Promise<void> {
  if (!auth || !tenantId || !clientId) return;
  try {
    await AgenticTokenCacheInstance.RefreshObservabilityToken(
      clientId,
      tenantId,
      context,
      auth,
      [SUBSTRATE_SCOPE],
    );
  } catch (err) {
    logger.debug({ err }, 'Agent 365: observability token preload skipped');
  }
}
