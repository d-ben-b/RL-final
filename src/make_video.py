"""make_video.py — render a side-by-side comparison clip of two trained policies
driving the SAME scenes, to visualise the paper's central finding (similar
training reward, opposite deployed behaviour).

Top row = DQN (safe), bottom row = PPO (crashes). Both are evaluated under the
common base test config with identical reset seeds, so the traffic is the same
and only the policy differs.

Usage (from repo root):
    python src/make_video.py --env merge-v0 --episodes 4 --out logs/demo_merge.mp4
Then compress to a small file:
    ffmpeg -y -i logs/demo_merge.mp4 -vf scale=-2:480 -crf 30 -preset slow logs/demo_merge_small.mp4
"""
import os, sys, argparse
import numpy as np
import torch
import gymnasium as gym
import highway_env  # noqa: F401  (registers envs)
import imageio
from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.join(os.path.dirname(__file__)))
from agents.networks import AttentionActorCritic, MlpActorCritic, MlpQNetwork
from envs.reward_shaper import get_env_config

DEVICE = torch.device("cpu")


def build_net(algo, arch, obs_shape, n_actions):
    if algo == "ppo":
        return (AttentionActorCritic(obs_shape, n_actions, 4) if arch == "attention"
                else MlpActorCritic(obs_shape, n_actions)).to(DEVICE)
    return MlpQNetwork(obs_shape, n_actions).to(DEVICE)


def load_policy(model_path, algo, arch, obs_shape, n_actions):
    net = build_net(algo, arch, obs_shape, n_actions)
    net.load_state_dict(torch.load(model_path, map_location=DEVICE))
    net.eval()

    def act(obs):
        t = torch.Tensor(obs).unsqueeze(0).to(DEVICE)
        with torch.no_grad():
            if algo == "ppo":
                feat = (net.extract_features(t) if arch == "attention"
                        else net.feature_extractor(t))
                return torch.argmax(net.actor(feat), dim=1).item()
            return torch.argmax(net(t), dim=1).item()
    return act


def label_bar(width, text, color=(30, 30, 30)):
    bar = Image.new("RGB", (width, 26), color)
    d = ImageDraw.Draw(bar)
    try:
        font = ImageFont.truetype("arial.ttf", 16)
    except Exception:
        font = ImageFont.load_default()
    d.text((8, 4), text, fill=(255, 255, 255), font=font)
    return np.array(bar)


def overlay(frame, step, crashed):
    """Draw a survival-time counter, and CRASHED if applicable."""
    img = Image.fromarray(frame.copy())
    d = ImageDraw.Draw(img)
    try:
        small = ImageFont.truetype("arial.ttf", 13)
        big = ImageFont.truetype("arial.ttf", 20)
    except Exception:
        small = big = ImageFont.load_default()
    d.rectangle([(0, 0), (110, 18)], fill=(0, 0, 0))
    d.text((4, 2), f"step {step}", fill=(255, 255, 0), font=small)
    if crashed:
        d.text((10, img.height // 2 - 12), "CRASHED", fill=(255, 50, 50), font=big)
    return np.array(img)


def run_pair(env_id, duration, top, bot, seeds, step_cap=130, hold_after=25):
    """Render paired episodes. Each episode runs until both policies terminate,
    a per-episode step cap, or `hold_after` frames past the first termination
    (so a surviving agent is shown continuing, without a minute-long freeze)."""
    cfg = get_env_config(env_id, "base", duration=duration)
    e_top = gym.make(env_id, render_mode="rgb_array"); e_top.unwrapped.configure(cfg)
    e_bot = gym.make(env_id, render_mode="rgb_array"); e_bot.unwrapped.configure(cfg)

    o, _ = e_top.reset(seed=0)
    act_top = load_policy(top["model"], top["algo"], top["arch"], o.shape, e_top.action_space.n)
    act_bot = load_policy(bot["model"], bot["algo"], bot["arch"], o.shape, e_bot.action_space.n)

    frames = []
    for sd in seeds:
        ot, _ = e_top.reset(seed=sd)
        ob, _ = e_bot.reset(seed=sd)
        dt = db = False
        crashed_t = crashed_b = False
        last_t = e_top.render(); last_b = e_bot.render()
        step = 0; since_first_done = 0
        while step < step_cap:
            if not dt:
                a = act_top(ot)
                ot, _, term, trunc, it = e_top.step(a); last_t = e_top.render()
                if it.get("crashed", False): crashed_t = True
                dt = term or trunc
            if not db:
                a = act_bot(ob)
                ob, _, term, trunc, ib = e_bot.step(a); last_b = e_bot.render()
                if ib.get("crashed", False): crashed_b = True
                db = term or trunc

            ft = overlay(last_t, step + 1, crashed_t)
            fb = overlay(last_b, step + 1, crashed_b)
            w = max(ft.shape[1], fb.shape[1])
            frames.append(np.vstack([
                label_bar(w, top["label"], (20, 90, 20)), ft,
                label_bar(w, bot["label"], (90, 20, 20)), fb,
            ]))
            step += 1
            if dt or db:
                since_first_done += 1
                if (dt and db) or since_first_done >= hold_after:
                    break
        for _ in range(6):  # short end-of-episode hold (~0.6 s)
            frames.append(frames[-1])

    e_top.close(); e_bot.close()
    return frames


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--env", default="highway-v0")
    ap.add_argument("--duration", type=int, default=80)
    ap.add_argument("--episodes", type=int, default=3)
    ap.add_argument("--out", default="logs/demo.mp4")
    ap.add_argument("--fps", type=int, default=10)
    # which seeds to visualise (pick crashing PPO vs surviving DQN for contrast)
    ap.add_argument("--dqn_seed", type=int, default=3)
    ap.add_argument("--ppo_seed", type=int, default=1)
    args = ap.parse_args()

    es = args.env.replace("-v0", "").replace("-", "_")
    # merge: all seeds behave identically (DQN 0%, PPO 100%) -> seed 0 is fine
    if es == "merge":
        args.dqn_seed = args.ppo_seed = 0
    TOP = dict(model=f"models/custom_model_dqn_mlp_aggressive_{es}_dur80_seed{args.dqn_seed}.pth",
               algo="dqn", arch="mlp", label="DQN + MLP  (survives)")
    BOT = dict(model=f"models/custom_model_ppo_attention_aggressive_{es}_dur80_seed{args.ppo_seed}.pth",
               algo="ppo", arch="attention", label="PPO + Attention  (crashes)")

    print(f"Rendering {args.episodes} paired episodes on {args.env} "
          f"(DQN seed{args.dqn_seed} vs PPO seed{args.ppo_seed}) ...")
    frames = run_pair(args.env, args.duration, TOP, BOT, seeds=list(range(args.episodes)))
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    imageio.mimsave(args.out, frames, fps=args.fps, macro_block_size=1)
    print(f"Saved raw video: {args.out}  ({len(frames)} frames, ~{len(frames)/args.fps:.0f}s)")
