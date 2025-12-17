#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
NSFW 视频内容检测脚本
使用 NudeNet 模型检测视频中的敏感内容
输出: JSON 格式的检测结果
"""

import sys
import os
import json
import tempfile
import subprocess
import argparse
from pathlib import Path

def check_dependencies():
    """检查并安装依赖"""
    try:
        from nudenet import NudeDetector
        return True
    except ImportError:
        print("正在安装 NudeNet...", file=sys.stderr)
        subprocess.check_call([sys.executable, "-m", "pip", "install", "nudenet", "-q"])
        return True

def extract_frames(video_path: str, frame_count: int = 5) -> list:
    """使用 ffmpeg 从视频中提取关键帧"""
    temp_dir = tempfile.mkdtemp(prefix="nsfw_detect_")
    frames = []
    
    try:
        # 获取视频时长
        result = subprocess.run(
            ["ffprobe", "-v", "quiet", "-print_format", "json", "-show_format", video_path],
            capture_output=True, text=True
        )
        info = json.loads(result.stdout)
        duration = float(info.get("format", {}).get("duration", 60))
        
        # 计算帧提取时间点（均匀分布，避开开头和结尾）
        interval = duration / (frame_count + 1)
        
        for i in range(1, frame_count + 1):
            timestamp = interval * i
            frame_path = os.path.join(temp_dir, f"frame_{i}.jpg")
            
            subprocess.run(
                ["ffmpeg", "-ss", str(timestamp), "-i", video_path, 
                 "-vframes", "1", "-q:v", "2", "-y", frame_path],
                capture_output=True
            )
            
            if os.path.exists(frame_path):
                frames.append(frame_path)
    except Exception as e:
        print(f"帧提取错误: {e}", file=sys.stderr)
    
    return frames, temp_dir

def analyze_frames(frames: list, detector) -> dict:
    """分析所有帧并汇总结果"""
    all_detections = []
    nsfw_score = 0
    max_score = 0
    nsfw_classes = set()
    
    # 定义 NSFW 类别
    explicit_classes = {
        "FEMALE_BREAST_EXPOSED", "FEMALE_GENITALIA_EXPOSED",
        "MALE_GENITALIA_EXPOSED", "BUTTOCKS_EXPOSED",
        "ANUS_EXPOSED", "FEMALE_BREAST_COVERED",
        "exposed_breast", "exposed_genitalia", "exposed_buttocks",
        "exposed_anus", "exposed_belly"
    }
    
    for frame_path in frames:
        try:
            detections = detector.detect(frame_path)
            
            for det in detections:
                class_name = det.get("class", "")
                score = det.get("score", 0)
                
                all_detections.append({
                    "class": class_name,
                    "score": score
                })
                
                # 检查是否为 NSFW 内容
                if class_name in explicit_classes or "exposed" in class_name.lower():
                    nsfw_classes.add(class_name)
                    if score > max_score:
                        max_score = score
                    nsfw_score += score
        except Exception as e:
            print(f"分析帧错误: {e}", file=sys.stderr)
    
    # 计算平均 NSFW 分数
    avg_score = nsfw_score / len(frames) if frames else 0
    
    return {
        "is_nsfw": max_score > 0.5 or avg_score > 0.3,
        "max_score": round(max_score, 4),
        "avg_score": round(avg_score, 4),
        "nsfw_classes": list(nsfw_classes),
        "detection_count": len([d for d in all_detections if d["score"] > 0.5]),
        "frames_analyzed": len(frames)
    }

def cleanup(temp_dir: str, frames: list):
    """清理临时文件"""
    for frame in frames:
        try:
            os.remove(frame)
        except:
            pass
    try:
        os.rmdir(temp_dir)
    except:
        pass

def main():
    parser = argparse.ArgumentParser(description="NSFW 视频内容检测")
    parser.add_argument("video_path", help="视频文件路径")
    parser.add_argument("-f", "--frames", type=int, default=5, help="提取帧数 (默认: 5)")
    parser.add_argument("-t", "--threshold", type=float, default=0.5, help="NSFW 阈值 (默认: 0.5)")
    parser.add_argument("-v", "--verbose", action="store_true", help="详细输出")
    
    args = parser.parse_args()
    
    # 检查视频文件
    if not os.path.exists(args.video_path):
        result = {"error": "视频文件不存在", "is_nsfw": False}
        print(json.dumps(result, ensure_ascii=False))
        sys.exit(1)
    
    # 检查依赖
    check_dependencies()
    
    # 导入检测器
    from nudenet import NudeDetector
    
    if args.verbose:
        print(f"加载 NudeNet 模型...", file=sys.stderr)
    
    detector = NudeDetector()
    
    if args.verbose:
        print(f"提取视频帧: {args.video_path}", file=sys.stderr)
    
    # 提取帧
    frames, temp_dir = extract_frames(args.video_path, args.frames)
    
    if not frames:
        result = {"error": "无法提取视频帧", "is_nsfw": False}
        print(json.dumps(result, ensure_ascii=False))
        cleanup(temp_dir, frames)
        sys.exit(1)
    
    if args.verbose:
        print(f"分析 {len(frames)} 帧...", file=sys.stderr)
    
    # 分析帧
    result = analyze_frames(frames, detector)
    result["video_path"] = args.video_path
    result["threshold"] = args.threshold
    
    # 根据阈值调整判断
    if result["max_score"] < args.threshold and result["avg_score"] < args.threshold * 0.6:
        result["is_nsfw"] = False
    
    # 输出 JSON 结果
    print(json.dumps(result, ensure_ascii=False))
    
    # 清理
    cleanup(temp_dir, frames)
    
    # 返回码: 0=非NSFW, 1=NSFW
    sys.exit(1 if result["is_nsfw"] else 0)

if __name__ == "__main__":
    main()
